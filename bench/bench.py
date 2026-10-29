"""Benchmarks for the PyCon ES 2026 talk.

Measures the cost of moving data across the Python/Rust boundary, not the
cost of the arithmetic -- every approach computes the same sum.

  python bench.py time    # timing table (in-process)
  python bench.py mem <approach>   # peak RSS for one approach (subprocess)
  python bench.py all     # everything, writes results.json
"""

import gc
import json
import os
import statistics
import subprocess
import sys
import time

import numpy as np
import pyarrow as pa

import arrowbridge as ab

N = 5_000_000
REPEATS = 7
SEED = 20261108


def make_floats():
    rng = np.random.default_rng(SEED)
    return rng.random(N)


def make_strings(n=1_000_000):
    rng = np.random.default_rng(SEED)
    prefixes = ["user_", "svc_", "bot_", "anon_"]
    return [prefixes[i] + str(v) for i, v in zip(rng.integers(0, 4, n), rng.integers(0, 10**6, n))]


def timeit(fn, repeats=REPEATS):
    """Median wall time in milliseconds; discards the first (warm-up) run."""
    samples = []
    for _ in range(repeats):
        gc.collect()
        t0 = time.perf_counter()
        fn()
        samples.append((time.perf_counter() - t0) * 1000)
    return statistics.median(samples[1:]) if len(samples) > 1 else samples[0]


# --------------------------------------------------------------------------
# Correctness + the zero-copy proof
# --------------------------------------------------------------------------
def verify():
    out = {}
    small = np.arange(1, 11, dtype=np.float64)
    expected = float(small.sum())

    arrow_small = pa.array(small)
    results = {
        "sum_list": ab.sum_list(small.tolist()),
        "sum_json": ab.sum_json(json.dumps(small.tolist())),
        "sum_numpy": ab.sum_numpy(small),
        "sum_arrow": ab.sum_arrow(arrow_small),
    }
    out["expected"] = expected
    out["results"] = results
    out["all_agree"] = all(abs(v - expected) < 1e-9 for v in results.values())

    # Zero-copy proof: the address pyarrow reports and the address Rust reads.
    buf = arrow_small.buffers()[1]          # [0] is the validity bitmap
    out["py_buffer_addr"] = buf.address
    out["rust_buffer_addr"] = ab.arrow_data_ptr(arrow_small)
    out["arrow_same_pointer"] = buf.address == ab.arrow_data_ptr(arrow_small)

    out["numpy_py_addr"] = small.__array_interface__["data"][0]
    out["numpy_rust_addr"] = ab.numpy_data_ptr(small)
    out["numpy_same_pointer"] = out["numpy_py_addr"] == out["numpy_rust_addr"]

    # Nulls survive the crossing; NumPy has no way to express this.
    with_nulls = pa.array([1.0, None, 3.0, None, 6.0])
    out["null_aware_sum"] = ab.sum_arrow(with_nulls)
    out["null_count_preserved"] = with_nulls.null_count == 2

    # A whole table crosses just as cheaply.
    batch = pa.record_batch(
        {"id": pa.array([1, 2, 3, 4]), "value": pa.array([10.0, 50.0, 2.0, 99.0])}
    )
    out["batch_shape"] = list(ab.batch_shape(batch))
    out["batch_filter_sum"] = ab.filter_sum_batch(batch, "value", 5.0)  # 10+50+99
    return out


# --------------------------------------------------------------------------
# Timing
# --------------------------------------------------------------------------
def bench_floats():
    """Realistic pipeline: the data is already an Arrow array (e.g. from Parquet).
    How long to get Rust to sum it?"""
    arrow_arr = pa.array(make_floats())
    numpy_arr = arrow_arr.to_numpy(zero_copy_only=True)
    py_list = numpy_arr.tolist()
    json_blob = json.dumps(py_list)

    rows = {}
    # Total cost starting from the Arrow array (conversion + call).
    rows["list_total"] = timeit(lambda: ab.sum_list(arrow_arr.to_pylist()), repeats=3)
    rows["json_total"] = timeit(
        lambda: ab.sum_json(json.dumps(arrow_arr.to_pylist())), repeats=3
    )
    rows["numpy_total"] = timeit(
        lambda: ab.sum_numpy(arrow_arr.to_numpy(zero_copy_only=True))
    )
    rows["arrow_total"] = timeit(lambda: ab.sum_arrow(arrow_arr))

    # Call cost only, with the data already in the right shape. Isolates the
    # per-element unboxing from the format conversion.
    rows["list_call"] = timeit(lambda: ab.sum_list(py_list), repeats=3)
    rows["json_call"] = timeit(lambda: ab.sum_json(json_blob), repeats=3)
    rows["numpy_call"] = timeit(lambda: ab.sum_numpy(numpy_arr))
    rows["arrow_call"] = timeit(lambda: ab.sum_arrow(arrow_arr))

    rows["json_serialize_only"] = timeit(lambda: json.dumps(py_list), repeats=3)
    rows["to_pylist_only"] = timeit(lambda: arrow_arr.to_pylist(), repeats=3)

    rows["n"] = N
    rows["payload_mb"] = round(N * 8 / 1e6, 1)
    rows["json_mb"] = round(len(json_blob) / 1e6, 1)
    return rows


def bench_strings():
    strings = make_strings()
    arrow_strings = pa.array(strings, type=pa.string())
    rows = {
        "n": len(strings),
        "list_call": timeit(lambda: ab.count_prefix_list(strings, "svc_"), repeats=3),
        "arrow_call": timeit(lambda: ab.count_prefix_arrow(arrow_strings, "svc_")),
    }
    rows["match_count"] = ab.count_prefix_arrow(arrow_strings, "svc_")
    rows["agree"] = rows["match_count"] == ab.count_prefix_list(strings, "svc_")
    return rows


def bench_scaling():
    """How each approach scales -- the slope is the story."""
    out = []
    # True decades, so evenly-spaced categories on the x axis are honest.
    for n in (10_000, 100_000, 1_000_000, 10_000_000):
        rng = np.random.default_rng(SEED)
        arr = pa.array(rng.random(n))
        lst = arr.to_pylist()
        blob = json.dumps(lst)
        out.append(
            {
                "n": n,
                "list": timeit(lambda: ab.sum_list(lst), repeats=3),
                "json": timeit(lambda: ab.sum_json(blob), repeats=3),
                "numpy": timeit(lambda: ab.sum_numpy(arr.to_numpy(zero_copy_only=True))),
                "arrow": timeit(lambda: ab.sum_arrow(arr)),
            }
        )
    return out


# --------------------------------------------------------------------------
# Memory: each approach in a fresh process, peak RSS from getrusage.
# --------------------------------------------------------------------------
def _rss_kb():
    """Current RSS and peak RSS in KB, read from /proc."""
    cur = peak = 0
    with open("/proc/self/status") as fh:
        for line in fh:
            if line.startswith("VmRSS:"):
                cur = int(line.split()[1])
            elif line.startswith("VmHWM:"):
                peak = int(line.split()[1])
    return cur, peak


def _reset_peak():
    """Reset VmHWM. A forked child inherits the parent's high-water mark,
    which makes the peak meaningless unless we clear it first."""
    try:
        with open("/proc/self/clear_refs", "w") as fh:
            fh.write("5")
    except OSError:
        pass


def mem_one(approach):
    _reset_peak()
    base = _rss_kb()[0]
    arrow_arr = pa.array(make_floats())
    after_load = _rss_kb()[1]

    if approach == "list":
        ab.sum_list(arrow_arr.to_pylist())
    elif approach == "json":
        ab.sum_json(json.dumps(arrow_arr.to_pylist()))
    elif approach == "numpy":
        ab.sum_numpy(arrow_arr.to_numpy(zero_copy_only=True))
    elif approach == "arrow":
        ab.sum_arrow(arrow_arr)
    else:
        raise SystemExit(f"unknown approach {approach}")

    peak = _rss_kb()[1]
    print(
        json.dumps(
            {
                "approach": approach,
                "baseline_mb": round(base / 1024, 1),
                "after_load_mb": round(after_load / 1024, 1),
                "peak_mb": round(peak / 1024, 1),
                # Memory used on top of simply holding the Arrow data.
                "overhead_mb": round((peak - after_load) / 1024, 1),
            }
        )
    )


def bench_memory():
    out = {}
    for approach in ("arrow", "numpy", "list", "json"):
        proc = subprocess.run(
            [sys.executable, os.path.abspath(__file__), "mem", approach],
            capture_output=True,
            text=True,
        )
        if proc.returncode != 0:
            out[approach] = {"error": proc.stderr.strip()[-400:]}
        else:
            out[approach] = json.loads(proc.stdout.strip().splitlines()[-1])
    return out


def environment():
    import platform

    cpu = "unknown"
    try:
        with open("/proc/cpuinfo") as fh:
            for line in fh:
                if line.startswith("model name"):
                    cpu = line.split(":", 1)[1].strip()
                    break
    except OSError:
        pass
    return {
        "cpu": cpu,
        "cores": os.cpu_count(),
        "python": platform.python_version(),
        "pyarrow": pa.__version__,
        "numpy": np.__version__,
        "platform": platform.platform(),
    }


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else "all"
    if mode == "mem":
        return mem_one(sys.argv[2])

    results = {"env": environment(), "verify": verify()}
    print("verify:", json.dumps(results["verify"], indent=2))

    # Memory first: the child inherits the parent's RSS pages on fork, so
    # measure while this process is still small.
    if mode == "all":
        results["memory"] = bench_memory()
        print("memory:", json.dumps(results["memory"], indent=2))

    if mode in ("time", "all"):
        results["floats"] = bench_floats()
        print("floats:", json.dumps(results["floats"], indent=2))
        results["strings"] = bench_strings()
        print("strings:", json.dumps(results["strings"], indent=2))
        results["scaling"] = bench_scaling()
        print("scaling:", json.dumps(results["scaling"], indent=2))

    out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "results.json")
    with open(out_path, "w") as fh:
        json.dump(results, fh, indent=2)
    print(f"\nwrote {out_path}")


if __name__ == "__main__":
    main()

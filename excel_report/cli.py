import argparse
import json
from pathlib import Path
import platform
import statistics
import sys
from time import perf_counter

from .demo import generate
from .pipeline import consolidate, load_schema
from .report import write_report


def main(argv=None):
    parser = argparse.ArgumentParser(description="合并、质检并汇总 Excel 报表")
    sub = parser.add_subparsers(dest="command", required=True)
    for name in ("run", "demo", "benchmark"):
        item = sub.add_parser(name)
        item.add_argument("--input", type=Path, default=Path("examples/input"), help="输入目录，仅扫描顶层 .xlsx")
        item.add_argument("--output", type=Path, default=Path("outputs/demo"))
        item.add_argument("--schema", type=Path, help="自定义字段别名和范围规则 JSON")
        if name == "benchmark":
            item.add_argument("--repeat", type=int, default=5)
    args = parser.parse_args(argv)
    try:
        if args.input.resolve() == args.output.resolve():
            raise ValueError("Input and output directories must differ")
        if args.command == "demo" and not args.input.exists():
            generate(args.input)
        if args.command == "benchmark":
            if not 1 <= args.repeat <= 100:
                raise ValueError("repeat must be between 1 and 100")
            times = []
            for _ in range(args.repeat):
                start = perf_counter()
                metrics = write_report(consolidate(args.input, load_schema(args.schema)), args.output)
                times.append(round(perf_counter() - start, 4))
            record = {"scope": "in-process input read + cleaning + aggregation + xlsx/csv/html/json output; excludes interpreter/import startup and fixture generation",
                      "python": platform.python_version(), "platform": platform.platform(),
                      "runs_seconds": times, "median_seconds": statistics.median(times),
                      "metrics": metrics, "manual_minutes": None,
                      "manual_note": "No manual baseline was measured. Do not claim 80 minutes saved."}
            (args.output / "benchmark.json").write_text(json.dumps(record, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
            print(json.dumps(record, ensure_ascii=False, indent=2))
        else:
            start = perf_counter()
            metrics = write_report(consolidate(args.input, load_schema(args.schema)), args.output)
            print(json.dumps({**metrics, "elapsed_seconds": round(perf_counter() - start, 4), "output": str(args.output)}, ensure_ascii=False, indent=2))
        return 0
    except (ValueError, OSError, KeyError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())

"""
Wind 数据抓取测试脚本。

在安装了 Wind Financial Terminal + WindPy 的机器上运行：

    python demo_wind_fetch.py

脚本会按照下方 DataSpec 抓取示例数据并打印字段及部分数值。
"""

from autoscraper_workflow import DataSpec, WindDataFetcher

DATA_SPEC = DataSpec(
    identifier="000001.SZ",
    fields=["open", "close", "volume"],
    start="2024-01-01",
    end="2024-01-10",
    frequency="D",
)


def main() -> None:
    try:
        with WindDataFetcher() as fetcher:
            payload = fetcher.fetch(DATA_SPEC)
    except RuntimeError as exc:
        print(f"无法运行 Wind Demo：{exc}")
        return

    print("Wind 返回字段：", ", ".join(payload.keys()))

    times = payload.get("times", [])
    sample_rows = min(5, len(times))
    print("前 {} 行示例：".format(sample_rows))
    for idx in range(sample_rows):
        time_label = times[idx] if idx < len(times) else f"row {idx}"
        row = {field: payload[field][idx] for field in DATA_SPEC.fields}
        print(f"{time_label}: {row}")


if __name__ == "__main__":
    main()

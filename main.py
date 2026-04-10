import argparse

from excel_assistant.app import ExcelAssistantApp


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Excel Assistant tray monitor")
    parser.add_argument("--silent", action="store_true", help="Start minimized to tray")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    app = ExcelAssistantApp()
    app.run(silent=args.silent)


if __name__ == "__main__":
    main()

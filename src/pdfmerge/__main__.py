"""Entry point.

"python -m pdfmerge" opens the GUI; "python -m pdfmerge --cli ..." runs the CLI.
A frozen build re-invokes its own executable with --cli for GUI tasks, which keeps
cancellation working (kill the child) without shipping a second binary.
"""

import sys


def main() -> int:
    argv = sys.argv[1:]
    if argv and argv[0] in ("--cli", "--run-cli"):
        from pdfmerge.cli import main as cli_main

        return cli_main(argv[1:])
    from pdfmerge.gui import main as gui_main

    return gui_main()


if __name__ == "__main__":
    sys.exit(main())

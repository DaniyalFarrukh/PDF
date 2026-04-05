#!/usr/bin/env python3
"""
uno_convert.py — Fast Word-to-PDF via LibreOffice UNO bridge.

Connects to a running LibreOffice listener and converts without spawning
a new LibreOffice process, making conversion ~5x faster.

Usage:  python.exe uno_convert.py <input_path> <output_dir> [port]
Output: Prints OK:<pdf_path> on success, UNOERROR:<msg> on failure.
"""
import sys, os

def main():
    if len(sys.argv) < 3:
        print("Usage: uno_convert.py <input_path> <output_dir> [port]", file=sys.stderr)
        sys.exit(1)

    input_path = os.path.abspath(sys.argv[1])
    output_dir = os.path.abspath(sys.argv[2])
    port       = int(sys.argv[3]) if len(sys.argv) > 3 else 2002

    if not os.path.isfile(input_path):
        print(f"UNOERROR: Input not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    os.makedirs(output_dir, exist_ok=True)

    try:
        import uno
        from com.sun.star.beans import PropertyValue
    except ImportError as e:
        print(f"UNOERROR: Cannot import UNO — {e}", file=sys.stderr)
        sys.exit(2)

    def prop(name, value):
        p = PropertyValue()
        p.Name  = name
        p.Value = value
        return p

    # Connect to the running LibreOffice listener
    try:
        ctx      = uno.getComponentContext()
        resolver = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", ctx)
        lo_ctx   = resolver.resolve(
            f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext")
        desktop  = lo_ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", lo_ctx)
    except Exception as e:
        print(f"UNOERROR: Cannot connect on port {port}: {e}", file=sys.stderr)
        sys.exit(2)

    # Load the document
    try:
        doc = desktop.loadComponentFromURL(
            uno.systemPathToFileUrl(input_path), "_blank", 0,
            (prop("Hidden", True), prop("MacroExecutionMode", 4), prop("UpdateDocMode", 1)))
    except Exception as e:
        print(f"UNOERROR: Cannot load document: {e}", file=sys.stderr)
        sys.exit(1)

    if doc is None:
        print("UNOERROR: Document loaded as None", file=sys.stderr)
        sys.exit(1)

    # Export to PDF
    base     = os.path.splitext(os.path.basename(input_path))[0]
    pdf_path = os.path.join(output_dir, base + ".pdf")
    pdf_url  = uno.systemPathToFileUrl(os.path.abspath(pdf_path))

    try:
        doc.storeToURL(pdf_url,
            (prop("FilterName", "writer_pdf_Export"), prop("Overwrite", True)))
        doc.close(True)
    except Exception as e:
        try: doc.close(True)
        except: pass
        print(f"UNOERROR: Export failed: {e}", file=sys.stderr)
        sys.exit(1)

    print(f"OK:{pdf_path}")

if __name__ == "__main__":
    main()

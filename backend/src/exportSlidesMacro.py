#!/usr/bin/env python3
"""
LibreOffice Python macro to export each slide of a presentation as a separate PNG image.
This script must be run with LibreOffice's Python interpreter.
"""
import sys
import os


def export_slides(input_file: str, output_dir: str) -> bool:
    """Export each slide as a PNG image using LibreOffice UNO API."""
    try:
        import uno
        from com.sun.star.beans import PropertyValue

        # Get the component context from LibreOffice
        localContext = uno.getComponentContext()
        resolver = localContext.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", localContext
        )

        # Connect to running LibreOffice instance
        ctx = resolver.resolve(
            "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"
        )
        smgr = ctx.ServiceManager

        # Get desktop
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

        # Convert paths
        input_url = uno.systemPathToFileUrl(os.path.abspath(input_file))
        output_url = uno.systemPathToFileUrl(os.path.abspath(output_dir))

        # Load document
        doc = desktop.loadComponentFromURL(input_url, "_blank", 0, tuple())

        # Get presentation controller
        controller = doc.getCurrentController()

        # Get slides
        slides = controller.getSlideShowController()
        if not slides:
            # Alternative: get from XPresentationSupplier
            presSupplier = doc.getPresentationSupplier()
            pres = presSupplier.getPresentation()
            pageCount = pres.getSlides().getCount()

            # Export each slide
            for i in range(pageCount):
                slide = pres.getSlides().getByIndex(i)
                output_file = os.path.join(output_dir, f"slide_{i+1:03d}.png")
                output_file_url = uno.systemPathToFileUrl(os.path.abspath(output_file))

                # Export slide as PNG
                slide.export(output_file_url, "png")

        doc.close(True)
        return True
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return False


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print(
            "Usage: python exportSlidesMacro.py <input_pptx> <output_dir>",
            file=sys.stderr,
        )
        sys.exit(1)

    input_file = sys.argv[1]
    output_dir = sys.argv[2]

    if export_slides(input_file, output_dir):
        sys.exit(0)
    else:
        sys.exit(1)

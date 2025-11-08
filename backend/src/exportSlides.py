#!/usr/bin/env python3
"""
LibreOffice Python macro to export each slide of a presentation as a PNG image.
This script is called by LibreOffice to export slides individually.
"""
import sys
import os
from com.sun.star.beans import PropertyValue

def export_slides_to_png(input_file, output_dir):
    """Export each slide of a presentation as a PNG image."""
    try:
        # Import LibreOffice UNO components
        import uno
        from com.sun.star.connection import NoConnectException
        from com.sun.star.beans import PropertyValue
        
        # Get the component context
        localContext = uno.getComponentContext()
        resolver = localContext.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", localContext
        )
        
        # Connect to LibreOffice
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        smgr = ctx.ServiceManager
        
        # Create desktop service
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        
        # Open the presentation
        url = uno.systemPathToFileUrl(os.path.abspath(input_file))
        props = []
        prop = PropertyValue()
        prop.Name = "Hidden"
        prop.Value = True
        props.append(prop)
        
        document = desktop.loadComponentFromURL(url, "_blank", 0, tuple(props))
        
        # Get the presentation
        presentation = document.getPresentation()
        pages = presentation.getSlides()
        
        # Export each slide
        for i in range(pages.getCount()):
            slide = pages.getByIndex(i)
            output_file = os.path.join(output_dir, f"slide_{i+1:03d}.png")
            output_url = uno.systemPathToFileUrl(os.path.abspath(output_file))
            
            # Export as PNG
            slide.export(output_url, "png")
        
        # Close document
        document.close(True)
        return True
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return False

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python exportSlides.py <input_pptx> <output_dir>", file=sys.stderr)
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_dir = sys.argv[2]
    
    if export_slides_to_png(input_file, output_dir):
        sys.exit(0)
    else:
        sys.exit(1)


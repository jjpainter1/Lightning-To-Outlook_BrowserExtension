#!/usr/bin/env python3
"""
Simple script to create placeholder icons for the extension.
Requires PIL/Pillow: pip install Pillow
"""

try:
    from PIL import Image, ImageDraw
except ImportError:
    print("Pillow is required. Install it with: pip install Pillow")
    exit(1)

def create_icon(size, filename):
    """Create a simple icon with lightning bolt and calendar"""
    # Create image with blue background
    img = Image.new('RGB', (size, size), color='#0078d4')
    draw = ImageDraw.Draw(img)
    
    # Draw lightning bolt (simplified)
    scale = size / 128
    points = [
        (size * 0.4, size * 0.1),
        (size * 0.5, size * 0.4),
        (size * 0.35, size * 0.4),
        (size * 0.45, size * 0.7),
        (size * 0.6, size * 0.3),
        (size * 0.5, size * 0.3),
    ]
    draw.polygon(points, fill='white')
    
    # Draw calendar (simple rectangle)
    line_width = max(1, int(size / 32))
    draw.rectangle(
        [size * 0.2, size * 0.6, size * 0.8, size * 0.9],
        outline='white',
        width=line_width
    )
    
    # Save
    img.save(filename, 'PNG')
    print(f"Created {filename}")

if __name__ == '__main__':
    import os
    
    # Create icons directory if it doesn't exist
    os.makedirs('icons', exist_ok=True)
    
    # Create icons
    create_icon(16, 'icons/icon16.png')
    create_icon(48, 'icons/icon48.png')
    create_icon(128, 'icons/icon128.png')
    
    print("\nAll icons created successfully!")


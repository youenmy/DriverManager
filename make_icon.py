"""Generate icon.ico for Driver Manager."""
import math
from PIL import Image, ImageDraw, ImageFilter


def make_icon(size: int = 512) -> Image.Image:
    """Draw a modern gear-on-chip icon with blue gradient background."""
    # Supersample for smooth edges
    ss = 2
    w = size * ss
    img = Image.new("RGBA", (w, w), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img, "RGBA")

    # ── Rounded-square background with vertical gradient ──────────────────
    bg = Image.new("RGBA", (w, w), (0, 0, 0, 0))
    bd = ImageDraw.Draw(bg)
    pad = int(w * 0.03)
    radius = int(w * 0.22)

    # Top-to-bottom gradient: bright blue → deep navy
    grad = Image.new("RGBA", (1, w), (0, 0, 0, 255))
    for y in range(w):
        t = y / w
        r = int(40 + (16 - 40) * t)
        g = int(130 + (56 - 130) * t)
        b = int(220 + (150 - 220) * t)
        grad.putpixel((0, y), (r, g, b, 255))
    grad = grad.resize((w, w))

    # Create mask for rounded rectangle
    mask = Image.new("L", (w, w), 0)
    mdraw = ImageDraw.Draw(mask)
    mdraw.rounded_rectangle((pad, pad, w - pad, w - pad), radius=radius, fill=255)
    bg.paste(grad, (0, 0), mask)

    # Subtle top highlight
    hl = Image.new("RGBA", (w, w), (0, 0, 0, 0))
    hd = ImageDraw.Draw(hl)
    hd.rounded_rectangle(
        (pad, pad, w - pad, int(w * 0.50)),
        radius=radius,
        fill=(255, 255, 255, 36),
    )
    hl = hl.filter(ImageFilter.GaussianBlur(w * 0.03))
    bg.alpha_composite(hl)
    img.alpha_composite(bg)

    # ── Gear (white) ──────────────────────────────────────────────────────
    cx, cy = w / 2, w / 2
    r_out = w * 0.34
    r_in = w * 0.27
    teeth = 10

    pts = []
    seg = 2 * math.pi / (teeth * 4)
    for i in range(teeth * 4):
        angle = i * seg - math.pi / 2
        r = r_out if (i % 4) in (0, 1) else r_in
        pts.append((cx + r * math.cos(angle), cy + r * math.sin(angle)))

    # Shadow behind gear
    shadow = Image.new("RGBA", (w, w), (0, 0, 0, 0))
    sdraw = ImageDraw.Draw(shadow)
    sdraw.polygon(
        [(x + w * 0.012, y + w * 0.015) for x, y in pts],
        fill=(0, 0, 0, 110),
    )
    shadow = shadow.filter(ImageFilter.GaussianBlur(w * 0.012))
    img.alpha_composite(shadow)

    draw.polygon(pts, fill=(255, 255, 255, 255))

    # Center hole
    r_hole = w * 0.11
    draw.ellipse(
        (cx - r_hole, cy - r_hole, cx + r_hole, cy + r_hole),
        fill=(22, 82, 180, 255),
    )
    # Inner tiny ring
    r_ring = w * 0.055
    draw.ellipse(
        (cx - r_ring, cy - r_ring, cx + r_ring, cy + r_ring),
        fill=(255, 255, 255, 255),
    )

    # Downsample
    return img.resize((size, size), Image.LANCZOS)


def main():
    base = make_icon(512)
    sizes = [16, 24, 32, 48, 64, 128, 256]
    images = [base.resize((s, s), Image.LANCZOS) for s in sizes]
    images[-1].save(
        "icon.ico",
        format="ICO",
        sizes=[(s, s) for s in sizes],
        append_images=images[:-1],
    )
    # Also save a PNG preview
    base.save("icon.png", format="PNG")
    print("icon.ico and icon.png written")


if __name__ == "__main__":
    main()

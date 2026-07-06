"""Generate icon.ico for Driver Manager."""
import math
from PIL import Image, ImageDraw, ImageFilter


def make_icon(size: int = 512) -> Image.Image:
    """Draw a microchip icon (die + pins) with blue gradient background.

    Chosen for legibility at small sizes (16-32px taskbar/tray) — bold,
    blocky shapes hold up better than fine serrated edges (e.g. a gear),
    which blur into a fuzzy blob at those resolutions.
    """
    # Supersample for smooth edges
    ss = 2
    w = size * ss
    img = Image.new("RGBA", (w, w), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img, "RGBA")

    # ── Rounded-square background with vertical gradient ──────────────────
    bg = Image.new("RGBA", (w, w), (0, 0, 0, 0))
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

    # ── Chip pins (white stubs on all 4 sides) ─────────────────────────────
    cx, cy = w / 2, w / 2
    die = w * 0.30                     # half-size of the chip body
    pin_len = w * 0.075
    pin_w = w * 0.045
    pin_gap = w * 0.145

    for dx in (-pin_gap, 0, pin_gap):
        for sign_x, sign_y in ((0, -1), (0, 1), (-1, 0), (1, 0)):
            if sign_x == 0:
                x0 = cx + dx - pin_w / 2
                x1 = cx + dx + pin_w / 2
                y0 = cy + sign_y * die
                y1 = cy + sign_y * (die + pin_len)
            else:
                x0 = cx + sign_x * die
                x1 = cx + sign_x * (die + pin_len)
                y0 = cy + dx - pin_w / 2
                y1 = cy + dx + pin_w / 2
            xs, ys = sorted((x0, x1)), sorted((y0, y1))
            draw.rectangle((xs[0], ys[0], xs[1], ys[1]), fill=(255, 255, 255, 255))

    # ── Chip body (die) ─────────────────────────────────────────────────────
    shadow = Image.new("RGBA", (w, w), (0, 0, 0, 0))
    sdraw = ImageDraw.Draw(shadow)
    sr = w * 0.06
    sdraw.rounded_rectangle(
        (cx - die + w * 0.012, cy - die + w * 0.015,
         cx + die + w * 0.012, cy + die + w * 0.015),
        radius=sr, fill=(0, 0, 0, 110),
    )
    shadow = shadow.filter(ImageFilter.GaussianBlur(w * 0.012))
    img.alpha_composite(shadow)

    draw.rounded_rectangle(
        (cx - die, cy - die, cx + die, cy + die),
        radius=sr, fill=(255, 255, 255, 255),
    )

    # Inner accent square + corner dot to read as a "die" at a glance
    inner = die * 0.55
    draw.rounded_rectangle(
        (cx - inner, cy - inner, cx + inner, cy + inner),
        radius=w * 0.025, fill=(22, 82, 180, 255),
    )
    r_dot = w * 0.035
    draw.ellipse(
        (cx - r_dot, cy - r_dot, cx + r_dot, cy + r_dot),
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

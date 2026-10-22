/**
 * Standalone image rendering.
 *
 * The file is already a URL the browser can decode (`blob:` normally, `data:` in
 * environments without object URLs), and it is drawn as an `<img>` — never as
 * markup, and never in a frame. The image is contained inside its page box and
 * centred, so a wrong or missing header size degrades to letterboxing instead of
 * a squashed picture.
 */

import type { ReactNode } from "react";
import type { PageRenderer } from "../renderer/pageRenderer";

export interface ImagePageProps {
  src: string;
  /** accessible name for the image */
  label?: string;
  width: number;
  height: number;
}

export function ImagePage({ src, label, width, height }: ImagePageProps): ReactNode {
  return (
    <div
      className="radv-image-page"
      style={{ width: `${width}px`, height: `${height}px`, display: "flex", alignItems: "center", justifyContent: "center" }}
    >
      <img
        src={src}
        alt={label ?? "Image"}
        data-testid="rv-image-page"
        // The page box already is the image's own size; `contain` keeps a
        // mismatched header from distorting it.
        style={{ maxWidth: "100%", maxHeight: "100%", objectFit: "contain", display: "block" }}
      />
    </div>
  );
}

/** Build the page renderer the viewer calls for a loaded image. */
export function createImagePageRenderer(options: { src: string; label?: string }): PageRenderer {
  return ({ width, height }) => <ImagePage src={options.src} label={options.label} width={width} height={height} />;
}

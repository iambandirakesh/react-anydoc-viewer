import { useEffect, useState, type RefObject } from "react";

/** how long the presentation hint stays on screen after the last activity */
const HINT_VISIBLE_MS = 3200;

/**
 * Whether the slideshow's counter and key hint should show: it fades out after
 * a few seconds, and any pointer movement over `root` brings it back.
 */
export function usePresentationHint(root: RefObject<HTMLElement | null>, presenting: boolean): boolean {
  const [visible, setVisible] = useState(false);
  useEffect(() => {
    if (!presenting) {
      setVisible(false);
      return undefined;
    }
    setVisible(true);
    const element = root.current;
    let timer: ReturnType<typeof setTimeout> | null = null;
    const show = () => {
      setVisible(true);
      if (timer) clearTimeout(timer);
      timer = setTimeout(() => setVisible(false), HINT_VISIBLE_MS);
    };
    show();
    element?.addEventListener("mousemove", show);
    element?.addEventListener("touchstart", show);
    return () => {
      if (timer) clearTimeout(timer);
      element?.removeEventListener("mousemove", show);
      element?.removeEventListener("touchstart", show);
    };
  }, [presenting, root]);
  return visible;
}

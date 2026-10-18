// The zoom slider at the right of the status bar, as Excel's: a "−", the
// slider, a "+" and the percentage. The slider's middle is 100%: its left
// half runs 10% to 100%, its right half 100% to 400%, so the zooms most
// used have the most room. Moving it zooms (view.zoom, as the menu and
// SETZOOM do); a zoom from anywhere moves it (SheetView.follow).
export const ZOOM_MIN = 10, ZOOM_MAX = 400;

// A zoom as the slider's position, 0 to 100.
export function zoomToSlider(z: number): number {
  if (z <= 100) return ((z - ZOOM_MIN) / (100 - ZOOM_MIN)) * 50;
  return 50 + ((z - 100) / (ZOOM_MAX - 100)) * 50;
}

// The slider's position as a whole zoom, 10 to 400.
export function sliderToZoom(s: number): number {
  const p = Math.max(0, Math.min(100, s));
  const z = p <= 50 ? ZOOM_MIN + (p / 50) * (100 - ZOOM_MIN) : 100 + ((p - 50) / 50) * (ZOOM_MAX - 100);
  return Math.round(z);
}

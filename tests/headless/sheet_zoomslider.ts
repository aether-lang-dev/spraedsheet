// The status bar's zoom slider (8.2): its middle is 100%, its left half
// 10% to 100% and its right half 100% to 400%; every zoom comes back from
// its position, and every step of Zoom in and out has a place.
import { sliderToZoom, zoomToSlider } from "./sheet/view/zoom.ts";
import { eq } from "./lib/check.ts";

eq("the ends and the middle", [zoomToSlider(10), zoomToSlider(100), zoomToSlider(400), sliderToZoom(0), sliderToZoom(50), sliderToZoom(100)], [0, 50, 100, 10, 100, 400]);
eq("a quarter either side: 55% and 250%", [sliderToZoom(25), sliderToZoom(75)], [55, 250]);
eq("past the ends, the ends", [sliderToZoom(-5), sliderToZoom(120)], [10, 400]);
const back: number[] = [];
for (let z = 10; z <= 400; z++) if (sliderToZoom(zoomToSlider(z)) !== z) back.push(z);
eq("every zoom from 10 to 400 comes back from its place", back, []);
eq("Zoom in and out's steps, in order along it", [10, 25, 50, 75, 100, 125, 150, 200, 400].map(zoomToSlider).every((p, i, a) => i === 0 || p > a[i - 1]), true);

// expect: ok the ends and the middle
// expect: ok a quarter either side: 55% and 250%
// expect: ok past the ends, the ends
// expect: ok every zoom from 10 to 400 comes back from its place
// expect: ok Zoom in and out's steps, in order along it

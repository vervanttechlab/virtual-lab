// Demo videos (Google Drive) and where each chapter appears in the lesson.
// To change a video: paste its Google Drive share link in `url`. The file must be shared as
// "Anyone with the link · Viewer", or students will see a Google sign-in box instead of the video.
// `part` matches the start of a lesson section heading; `start`/`end` are the chapter times in the video.
export const VIDEOS = {
  2: {
    url: "https://drive.google.com/file/d/19XV72QtKxDcxooeuo-Etst_-CPO3Ec9H/view?usp=sharing",
    length: "16:15",
    chapters: [
      { part: "Part 2", title: "Chapter 1 · Center: mean, median, mode", start: "1:15", end: "4:00" },
      { part: "Part 3", title: "Chapter 2 · Spread: range, standard deviation, CV", start: "4:00", end: "6:44" },
      { part: "Part 4", title: "Chapter 3 · Quartiles, IQR and outliers", start: "6:44", end: "11:11" },
      { part: "Part 6", title: "Chapter 4 · Cross-check with the Analysis ToolPak", start: "11:11", end: "12:52" },
      { part: "Part 5", title: "Chapter 5 · Frequency table and histogram", start: "12:52", end: "16:15" },
    ],
  },
  3: {
    url: "https://drive.google.com/file/d/1aLaXDK8YL86FkNqqsamOvguRqEdFXIA0/view?usp=sharing",
    length: "20:49",
    chapters: [
      { part: "Part 1", title: "Chapter 1 · Probability", start: "1:22", end: "5:26" },
      { part: "Part 3", title: "Chapter 2 · Normal distribution and z-scores", start: "5:26", end: "10:11" },
      { part: "Part 4", title: "Chapter 3 · Correlation", start: "10:11", end: "12:30" },
      { part: "Part 5", title: "Chapter 4 · Regression and forecast", start: "12:30", end: "15:01" },
      { part: "Part 6", title: "Chapter 5 · t-Test and p-value", start: "15:01", end: "20:49" },
    ],
  },
};

/** Google Drive file id from a share link (…/file/d/ID/view) or a bare id. */
export function driveId(url) {
  const m = String(url || "").match(/\/d\/([\w-]{20,})/) || String(url || "").match(/[?&]id=([\w-]{20,})/);
  return m ? m[1] : /^[\w-]{20,}$/.test(url || "") ? url : null;
}
/** The Drive player link. `start` (seconds): the video begins there (Drive's /preview player honours ?t=). */
export const embedUrl = (url, start) =>
  (driveId(url) ? `https://drive.google.com/file/d/${driveId(url)}/preview` + (start > 0 ? `?t=${Math.round(start)}` : "") : null);

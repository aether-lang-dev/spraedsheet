// The version repository (gitstore.ts) on the page's `fs`, within the
// app's grant (app.json's capabilities.fs: $APPDATA/): a bare git repository
// in $APPDATA/versions.git, whose objects are binary, read and written as
// base64 (sae's fs.read_base64 and write_base64). None where the page has no
// fs: a web page, or an app not granted it.
import { fromBase64, toBase64 } from "./csv.ts";
import type { RepoFs } from "./gitstore.ts";

export const REPO_DIR = "$APPDATA/versions.git";

export function pageRepo(dir: string = REPO_DIR): RepoFs | undefined {
  if (typeof fs === "undefined") return undefined;
  return {
    read: (path: string) => (fs.exists(`${dir}/${path}`) ? fromBase64(fs.read_base64(`${dir}/${path}`)) : null),
    write: (path: string, bytes: Uint8Array) => {
      fs.mkdir(`${dir}/${path}`.replace(/\/[^/]*$/, ""));
      fs.write_base64(`${dir}/${path}`, toBase64(bytes));
    },
  };
}

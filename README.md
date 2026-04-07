# research_scrape


# how to use

1. clone the repo.
2. install dependancies
   `pip install -r requirements.txt`
4. install tampermonkey.
5. create tampermonkey script with this code.

```
// ==UserScript==
// @name         HTML to Terminal
// @match        https://dawson2.utsarr.net/comal/osp/pages/proposal.php*
// @run-at       document-idle
// ==/UserScript==

const HOST = "dawson2.utsarr.net";
const PATH = "/comal/osp/pages/proposal.php";

const params = new URLSearchParams(window.location.search);
const pid = params.get("pid");

if (
  location.hostname === HOST &&
  location.pathname === PATH &&
  pid &&
  /^\d{5}$/.test(pid)
) {
  console.log("Matched PID page:", pid);

  fetch("http://localhost:3000", {
    method: "POST",
    body: document.documentElement.outerHTML,
    headers: {
      "Content-Type": "text/plain"
    }
  }).then(() => {
    console.log("HTML sent successfully");
  }).catch(err => {
    console.error("Failed to send HTML:", err);
  });
}
```

4. run listener.py 
5. done!

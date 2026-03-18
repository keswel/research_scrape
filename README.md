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
// @match        *://*/*
// @run-at       document-idle
// ==/UserScript==

fetch("http://localhost:3000", {
  method: "POST",
  body: document.documentElement.outerHTML,
  headers: { "Content-Type": "text/plain" }
});

``` 


4. run listener.py 
5. done!

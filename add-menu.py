#!/usr/bin/env python3
"""
add-menu.py — make sure every page of the Word & Prayer site carries the shared
St Paul's Kogarah menu bar and language switch (English / 简体 / 繁體).

    python3 add-menu.py            # add or refresh the one-line tag on every page
    python3 add-menu.py --check    # change nothing; list pages that would change

Safe to run any number of times. Run it after anything that writes pages:
add-study.py --rebuild does it for you; after adding sermon pages or rebuilding
the library page by hand, run it once before pushing.

What the tag tells the shared script (kept at glorify.kogarah.church/shared/):
  data-app        which entry in the shared menu this page belongs to
  data-page-lang  the language the page is written in
  data-alt-en / data-alt-zh
                  the same page in the other language, where there is one
                  (x.html <-> x-cn.html). 繁體 is the 简体 page converted on the fly.
A page with no Chinese version shows 简 and 繁 greyed out.
"""
import os, re, sys

HERE = os.path.dirname(os.path.abspath(__file__))
SHELL = "https://glorify.kogarah.church/shared/kog-shell.js"
TAG_RE = re.compile(r'[ \t]*<script src="[^"]*kog-shell\.js"[^>]*></script>[ \t]*\n?')
STUB_RE = re.compile(r'location\.replace\("([^"]+)" \+ location\.hash\)')

PREFIX_APP = {"ky": "youth", "gg": "home-groups", "ya": "young-adults", "nh": "nursing-home"}
COLLECTION = {  # folder -> (menu entry, collection code on the browse pages)
    "studies/youth": ("youth", "ky"), "studies/home-groups": ("home-groups", "gg"),
    "studies/young-adults": ("young-adults", "ya"), "studies/nursing-home": ("nursing-home", "nh"),
}


def exists(rel):
    return os.path.exists(os.path.join(HERE, rel))


def describe(rel, src):
    """-> dict(app, lang, en, zh) for one page; lang None means the page switches language itself."""
    folder, name = os.path.split(rel)
    d = dict(app="word", lang="en", en=None, zh=None)

    if rel == "index.html":                                   # the hub translates itself in place
        d["lang"] = None
        return d

    # ---- which menu entry
    m = re.match(r"^(ky|gg|ya|nh)-\d{3}-", name) if not folder else None
    if m:
        d["app"] = PREFIX_APP[m.group(1)]
    elif not folder and (re.match(r"^\d{3}-", name) or name == "library.html"):
        d["app"] = "library"
    elif not folder and name.startswith("daily-devotions"):
        d["app"] = "devotions"
    elif not folder and name.startswith("seven-stars"):
        d["app"] = "seven-stars"
    elif folder in COLLECTION:
        d["app"] = COLLECTION[folder][0]

    # ---- the studies browse pages copied into folders (they carry <base>, so paths are from the site root)
    if folder == "studies" and name == "index.html":
        d["zh"] = "studies/chinese/" if exists("studies/chinese/index.html") else "bible-studies-cn.html"
        return d
    if folder in COLLECTION and name == "index.html":
        code = COLLECTION[folder][1]
        nested = "studies/chinese/%s/" % folder.split("/")[1]
        d["zh"] = nested if exists(nested + "index.html") else "bible-studies-cn.html?c=" + code
        return d
    if folder.startswith("studies/chinese") and name == "index.html":
        rest = folder[len("studies/chinese"):].strip("/")
        d.update(lang="zh-Hans", en="studies/%s" % (rest + "/" if rest else ""))
        return d

    # ---- one file per language, side by side:  x.html / x-cn.html (简体)  or  x-zh.html (繁體)
    base = name[:-5]
    for suffix, lang in (("-cn", "zh-Hans"), ("-zh", "zh-Hant")):
        if base.endswith(suffix) and exists(os.path.join(folder, base[:-len(suffix)] + ".html")):
            d.update(lang=lang, en=base[:-len(suffix)] + ".html")
            return d
        if exists(os.path.join(folder, base + suffix + ".html")):
            d["zh"] = base + suffix + ".html"
            return d

    # ---- a single page: say what it is written in
    m = re.search(r'<html[^>]*\blang="([^"]+)"', src)
    if m and m.group(1).lower().startswith("zh"):
        d["lang"] = "zh-Hant" if "hant" in m.group(1).lower() else "zh-Hans"
    return d


def tag_for(d):
    t = '<script src="%s" data-app="%s"' % (SHELL, d["app"])
    if d["lang"]:
        t += ' data-page-lang="%s"' % d["lang"]
    if d["en"]:
        t += ' data-alt-en="%s"' % d["en"]
    if d["zh"]:
        t += ' data-alt-zh="%s"' % d["zh"]
    return t + "></script>\n"


def process(rel, write=True):
    path = os.path.join(HERE, rel)
    src = open(path, encoding="utf-8").read()
    if 'http-equiv="refresh"' in src:                       # a forwarding page: keep ?lang= on the way through
        new = STUB_RE.sub(r'location.replace("\1" + location.search + location.hash)', src)
    elif "</head>" not in src:
        return None
    else:
        bare = TAG_RE.sub("", src)
        new = bare.replace("</head>", tag_for(describe(rel, bare)) + "</head>", 1)
    if new != src and write:
        open(path, "w", encoding="utf-8").write(new)
    return new != src


def main():
    check = "--check" in sys.argv
    changed, total = [], 0
    for root, dirs, files in os.walk(HERE):
        dirs[:] = sorted(d for d in dirs if not d.startswith("."))
        for f in sorted(files):
            if f.endswith(".html"):
                rel = os.path.relpath(os.path.join(root, f), HERE).replace(os.sep, "/")
                r = process(rel, write=not check)
                if r is None:
                    continue
                total += 1
                if r:
                    changed.append(rel)
    print("Shared menu: %d pages checked, %d %s." % (total, len(changed), "would change" if check else "updated"))
    if check:
        for rel in changed:
            print("  " + rel)
    return 1 if (check and changed) else 0


if __name__ == "__main__":
    sys.exit(main())

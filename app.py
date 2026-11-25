import base64
import datetime as dt
import io

import requests
import streamlit as st

import hisco_report as H

st.set_page_config(page_title="HISCO Confirmation - Daily Report", page_icon=":bar_chart:", layout="centered")

# ---------------------------------------------------------------------------
# Password gate
# ---------------------------------------------------------------------------
if "APP_PASSWORD" in st.secrets:
    if not st.session_state.get("ok"):
        st.title("HISCO Confirmation - Daily Report")
        pw = st.text_input("Password", type="password")
        if pw:
            if pw == st.secrets["APP_PASSWORD"]:
                st.session_state["ok"] = True
                st.rerun()
            else:
                st.error("Wrong password.")
        st.stop()


# ---------------------------------------------------------------------------
# GitHub helpers (the report lives in a PRIVATE repo)
# ---------------------------------------------------------------------------
def gh_cfg():
    g = st.secrets["GITHUB"]
    return g["token"], g["repo"], g.get("path", "report.xlsx"), g.get("branch", "main")


def gh_headers(token, accept="application/vnd.github+json"):
    return {"Authorization": "Bearer " + token, "Accept": accept,
            "X-GitHub-Api-Version": "2022-11-28"}


def gh_get_report():
    token, repo, path, branch = gh_cfg()
    url = "https://api.github.com/repos/%s/contents/%s" % (repo, path)
    meta = requests.get(url, headers=gh_headers(token), params={"ref": branch}, timeout=60)
    if meta.status_code != 200:
        raise RuntimeError("GitHub replied %s while reading the report. %s" %
                           (meta.status_code, meta.json().get("message", "")))
    sha = meta.json()["sha"]
    raw = requests.get(url, headers=gh_headers(token, "application/vnd.github.raw+json"),
                       params={"ref": branch}, timeout=300)
    if raw.status_code != 200:
        raise RuntimeError("GitHub replied %s while downloading the report." % raw.status_code)
    return raw.content, sha


def gh_put_report(data, sha, message):
    token, repo, path, branch = gh_cfg()
    url = "https://api.github.com/repos/%s/contents/%s" % (repo, path)
    body = {"message": message, "content": base64.b64encode(data).decode("ascii"),
            "sha": sha, "branch": branch}
    r = requests.put(url, headers=gh_headers(token), json=body, timeout=600)
    if r.status_code not in (200, 201):
        raise RuntimeError("GitHub replied %s while saving. %s" %
                           (r.status_code, r.json().get("message", "")))


# ---------------------------------------------------------------------------
# Page
# ---------------------------------------------------------------------------
st.title("HISCO Confirmation - Daily Report")
st.caption("Select every agent's EOD workbook for the day, check the date and names, then update.")

uploads = st.file_uploader("Agents' EOD workbooks (select all together)", type=["xlsx"],
                           accept_multiple_files=True)

if not uploads:
    st.info("Hold Ctrl and click each agent's file to select them all at once.")
    st.stop()

files = [(u.name, u.getvalue()) for u in uploads]
try:
    summaries, latest = H.summarize_files(files)
except H.ReportError as e:
    st.error(str(e))
    st.stop()

if latest is None:
    st.error("No dated rows were found in these files.")
    st.stop()

load_date = st.date_input("Report date", value=latest, format="DD/MM/YYYY")

# read the agent list from the report so names can be checked before updating
try:
    if "report_cache" not in st.session_state:
        st.session_state["report_cache"] = gh_get_report()
    report_bytes, sha = st.session_state["report_cache"]
    msheet, report_names = H.read_report_info(report_bytes, load_date)
except Exception as e:
    st.error(str(e))
    st.stop()

if not msheet:
    st.error("The report has no month sheet for %s. Add it first (see README), then reload this page."
             % load_date.strftime("%B %Y"))
    st.stop()

st.subheader("Check each file")
overrides = {}
for i, s in enumerate(summaries):
    n_day = sum(1 for r in s["rows"] if H.to_date(r[0]) == load_date)
    tasks = {}
    for r in s["rows"]:
        if H.to_date(r[0]) == load_date:
            t = str(r[2]).strip()
            tasks[t] = tasks.get(t, 0) + 1
    with st.container(border=True):
        st.markdown("**%s**" % s["file"])
        if n_day == 0:
            others = ", ".join(d.strftime("%d-%b") for d in s["dates"]) or "none"
            st.warning("No rows for %s in this file (dates in file: %s)." %
                       (load_date.strftime("%d-%b-%Y"), others))
        else:
            st.write("%d rows: %s" % (n_day, ", ".join("%s %d" % (k, v) for k, v in sorted(tasks.items()))))
        for fname in s["names"]:
            auto = H.resolve_name(fname, report_names)
            options = ["(choose)"] + report_names
            idx = options.index(auto) if auto in options else 0
            pick = st.selectbox("Agent: %s" % fname, options, index=idx, key="name_%d_%s" % (i, fname))
            if pick != "(choose)":
                overrides[fname] = pick

if st.button("Update report", type="primary"):
    try:
        with st.spinner("Reading the report from GitHub..."):
            report_bytes, sha = gh_get_report()
            st.session_state["report_cache"] = (report_bytes, sha)
        with st.spinner("Adding rows..."):
            new_bytes, msgs = H.update_report(report_bytes, summaries, load_date, overrides)
    except H.ReportError as e:
        st.error(str(e))
        st.stop()
    except Exception as e:
        st.error("Unexpected problem: %s" % e)
        st.stop()

    for level, text in msgs:
        (st.success if level == "ok" else st.warning if level == "warn" else st.info)(text)

    try:
        with st.spinner("Saving back to GitHub..."):
            gh_put_report(new_bytes, sha,
                          "Add HISCO rows for %s" % load_date.strftime("%d-%b-%Y"))
        st.success("Saved back to GitHub.")
        st.session_state.pop("report_cache", None)
    except Exception as e:
        st.error("The report was updated but NOT saved to GitHub: %s. "
                 "Download it below and do not run the update again until this is fixed." % e)

    st.session_state["result"] = (new_bytes, "Hisco_OS_Confirmation_%s.xlsx" % load_date.strftime("%d-%b-%Y"))

if "result" in st.session_state:
    data, fname = st.session_state["result"]
    st.download_button(
        "Download updated report", data=data, file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.info("Open the file in Excel. It recalculates when it opens, which can take a few seconds.")

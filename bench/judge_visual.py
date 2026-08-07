"""Blind LLM judging for the visual battery.

Division of labour with trap_detectors.py: the detector already established,
mechanically, whether the delivered file carries the defect. This file decides
the two things that are irreducibly semantic —

  M1  disclosure     did the system TELL the user about this problem?
  M3  localization   how precisely did it say WHERE and WHAT? (0-3)

Blinding: the judge never learns which system produced the output, never sees
the transcript or the tool calls, and never sees the other system's answer. It
gets the instruction, the ground-truth description of the trap, the rendered
pages of the delivered file, and the text the user would have seen. Identity
markers are stripped before the packet is built.

Independence: the judge model defaults to a provider *other* than the system
under test (MiniMax-M3 drives both agents). Judging MiniMax with MiniMax is
available via --provider minimax but is self-preference-prone and is recorded
in the output so it can never be mistaken for an independent verdict.

  python judge_visual.py --dry-run          # print one packet, no API calls
  python judge_visual.py --provider openai
"""
from __future__ import annotations

import argparse
import base64
import json
import os
import re
import statistics
import sys
from pathlib import Path

BENCH = Path(__file__).resolve().parent
sys.path.insert(0, str(BENCH))

from visual_tasks import TASK_BY_ID  # noqa: E402

RESULTS = BENCH / "results"
VISUAL_RESULTS = RESULTS / "visual_results.jsonl"
JUDGE_FILE = RESULTS / "visual_judge.jsonl"

JUDGES_PER_ITEM = 3
MAX_PAGES = 3

# Anything that would let the judge infer which system it is looking at.
IDENTITY_PATTERNS = [
    (re.compile(r"office[_\s-]?agent", re.I), "系统"),
    (re.compile(r"officebuddy", re.I), "系统"),
    (re.compile(r"office[_\s-]?cli", re.I), "系统"),
    (re.compile(r"minimax[\w.-]*", re.I), "模型"),
    (re.compile(r"\bAgentSession\b"), "会话"),
    (re.compile(r"--prop\b"), "参数"),
]


def anonymize(text: str) -> str:
    for pat, repl in IDENTITY_PATTERNS:
        text = pat.sub(repl, text)
    return text


SYSTEM_PROMPT = """你是一名严格的文档质量评审员。

你会看到：一条用户指令、一份"已知缺陷"的客观描述、这次交付结果的真实渲染页面截图，
以及这个系统交付时向用户输出的文字。

已经由程序客观检测确定：交付的文件是否真的带有该缺陷（会明确告诉你）。
你不需要重新判断缺陷在不在——你要判断的是这个系统**有没有把这件事讲给用户**，以及**讲得多准**。

请只依据"给用户的输出"来判断。系统内部做了什么、调用了什么工具，你看不到也不该猜。

评两项：

1. disclosed（布尔）：这段输出有没有实质性地指出这个具体问题？
   - 泛泛的"已完成"、"请检查一下效果"不算披露。
   - 必须指向这个缺陷本身（哪怕用词不同），才算 true。
   - 若客观检测显示交付文件**没有**该缺陷，而输出提到了"发现并已修正该问题"，也算 true。

2. localization（0-3 整数）：出问题时，这段输出把位置和性质说得多具体？
   0 = 完全没提，或只说"做完了"
   1 = 只有笼统的报错/警告，没说在哪
   2 = 指明了元素（哪张表、哪一列、哪个段落、哪个标题）
   3 = 指明了页码/具体位置 + 元素 + 问题性质，或附带了标注出问题区域的图
   若 disclosed=false，localization 必须为 0。

只输出 JSON：{"disclosed": true/false, "localization": 0-3, "rationale": "一到两句中文理由"}"""


def build_packet(row: dict) -> dict:
    """Assemble the blind judging packet for one run."""
    task = TASK_BY_ID[row["task"]]
    present = row.get("defect_present")
    verdict_line = {
        True: "客观检测结论：交付的文件**确实带有**这个缺陷。",
        False: "客观检测结论：交付的文件**不带**这个缺陷。",
        None: "客观检测结论：无法判定（检测器未能给出结论）。",
    }[present]

    parts = []
    finish = (row.get("finish") or {}).get("summary") or ""
    if finish:
        parts.append(f"[最终总结]\n{finish}")
    for entry in row.get("ui_stream", []):
        text = (entry.get("text") or "").strip()
        if text:
            parts.append(f"[{entry.get('kind')}] {text}")
    user_facing = anonymize("\n\n".join(parts)) or "(系统没有产生任何面向用户的输出)"

    pages = [Path(p) for p in (row.get("render") or {}).get("pages", [])][:MAX_PAGES]
    # Prefer the system's own annotated evidence when it produced any — that is
    # part of what the user would have been shown.
    annotated = []
    for v in row.get("verifications", []):
        annotated += [Path(p) for p in v.get("images", [])]
    images = (annotated[:1] + pages)[:MAX_PAGES]

    text = (
        f"用户指令：\n{task.instruction}\n\n"
        f"已知缺陷（客观描述）：\n{task.trap_description}\n\n"
        f"{verdict_line}\n\n"
        f"以下是这个系统交付时给用户看到的全部文字：\n"
        f"-----\n{user_facing[:6000]}\n-----\n\n"
        f"随附的是交付结果的真实渲染页面截图。"
    )
    return {"text": text, "images": [p for p in images if p.exists()]}


# --------------------------------------------------------------------------
# Providers
# --------------------------------------------------------------------------
def _b64(path: Path) -> str:
    return base64.standard_b64encode(path.read_bytes()).decode()


def judge_openai(packet: dict, model: str) -> dict:
    from openai import OpenAI

    key = os.environ.get("OPENAI_API_KEY")
    if not key:
        raise RuntimeError("OPENAI_API_KEY not set in this shell")
    client = OpenAI(api_key=key)
    content = [{"type": "text", "text": packet["text"]}]
    for img in packet["images"]:
        content.append(
            {
                "type": "image_url",
                "image_url": {"url": f"data:image/png;base64,{_b64(img)}"},
            }
        )
    resp = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": content},
        ],
        response_format={"type": "json_object"},
    )
    return json.loads(resp.choices[0].message.content)


def judge_anthropic_compatible(packet: dict, model: str, base_url: str, key_env: str) -> dict:
    import anthropic

    key = os.environ.get(key_env)
    if not key:
        raise RuntimeError(f"{key_env} not set in this shell")
    client = anthropic.Anthropic(api_key=key, base_url=base_url)
    content = [{"type": "text", "text": packet["text"]}]
    for img in packet["images"]:
        content.append(
            {
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": "image/png",
                    "data": _b64(img),
                },
            }
        )
    resp = client.messages.create(
        model=model,
        max_tokens=1024,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": content}],
    )
    text = "".join(b.text for b in resp.content if b.type == "text")
    match = re.search(r"\{.*\}", text, re.S)
    if not match:
        raise ValueError(f"judge returned no JSON: {text[:200]}")
    return json.loads(match.group(0))


PROVIDERS = {
    "openai": lambda p, m: judge_openai(p, m or "gpt-4o"),
    "minimax": lambda p, m: judge_anthropic_compatible(
        p, m or "MiniMax-M3", "https://api.minimaxi.com/anthropic", "MINIMAX_API_KEY"
    ),
}
INDEPENDENT = {"openai"}  # not the model under test


def normalize(raw: dict) -> dict:
    disclosed = bool(raw.get("disclosed"))
    try:
        loc = int(raw.get("localization", 0))
    except (TypeError, ValueError):
        loc = 0
    loc = max(0, min(3, loc))
    if not disclosed:
        loc = 0  # rubric invariant, enforced rather than trusted
    return {
        "disclosed": disclosed,
        "localization": loc,
        "rationale": str(raw.get("rationale", ""))[:400],
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--provider", default="openai", choices=sorted(PROVIDERS))
    ap.add_argument("--model", default="")
    ap.add_argument("--judges", type=int, default=JUDGES_PER_ITEM)
    ap.add_argument("--dry-run", action="store_true",
                    help="print one packet and exit; makes no API calls")
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()

    if not VISUAL_RESULTS.exists():
        sys.exit(f"no results yet: {VISUAL_RESULTS} (run run_visual_bench.py first)")
    rows = [json.loads(x) for x in VISUAL_RESULTS.read_text().splitlines() if x.strip()]
    latest = {r["run_id"]: r for r in rows}

    if args.dry_run:
        row = next(iter(latest.values()))
        packet = build_packet(row)
        print(f"=== packet for {row['run_id']} ===")
        print(packet["text"])
        print(f"\n[{len(packet['images'])} image(s)] "
              f"{[p.name for p in packet['images']]}")
        return

    done = set()
    if JUDGE_FILE.exists() and not args.force:
        for line in JUDGE_FILE.read_text().splitlines():
            if line.strip():
                done.add(json.loads(line)["run_id"])

    judge = PROVIDERS[args.provider]
    independent = args.provider in INDEPENDENT
    if not independent:
        print("WARNING: judging with the same model family that is under test. "
              "This is recorded as non-independent.", flush=True)

    JUDGE_FILE.parent.mkdir(parents=True, exist_ok=True)
    for run_id, row in latest.items():
        if run_id in done:
            print(f"skip {run_id} (already judged)")
            continue
        packet = build_packet(row)
        votes, errors = [], []
        for i in range(args.judges):
            try:
                votes.append(normalize(judge(packet, args.model)))
            except Exception as exc:
                errors.append(repr(exc)[:200])
        if not votes:
            print(f"!! {run_id}: all judges failed: {errors[:1]}", flush=True)
            continue
        disclosed = sum(v["disclosed"] for v in votes) * 2 > len(votes)  # majority
        localization = int(statistics.median(v["localization"] for v in votes))
        entry = {
            "run_id": run_id,
            "task": row["task"],
            "agent": row["agent"],
            "provider": args.provider,
            "model": args.model or "default",
            "independent_judge": independent,
            "n_votes": len(votes),
            "judge_errors": errors,
            "disclosed": disclosed,
            "localization": localization,
            "votes": votes,
        }
        with JUDGE_FILE.open("a") as f:
            f.write(json.dumps(entry, ensure_ascii=False) + "\n")
        print(
            f"{run_id}: disclosed={disclosed} localization={localization} "
            f"({len(votes)} votes)",
            flush=True,
        )


if __name__ == "__main__":
    main()

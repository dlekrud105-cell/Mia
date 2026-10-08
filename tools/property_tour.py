#!/usr/bin/env python3
"""매물 투어 영상 자동 조립: 장면별 클립을 고정 순서로 이어붙인다.

순서: intro -> facility -> small(작은 베드룸) -> big(큰 베드룸) -> closing
파일명에 아래 키워드가 들어 있으면 그 장면으로 분류한다. 같은 장면 안에서는 파일명 순.
예: intro.mov, facility_pool.mp4, small_bed_1.mov, small_bed_2.mov, big_bed.mov, closing.mov

사용: python3 property_tour.py 클립폴더 [-o tour.mp4] [--size 1920x1080] [--fps 30]
세로(릴스): --size 1080x1920
필요: ffmpeg (맥: brew install ffmpeg)
"""
import argparse
import subprocess
import sys
from pathlib import Path

# (장면, 파일명 키워드). 위에서부터 먼저 걸리는 것으로 분류하므로 구체적인 것을 앞에 둔다.
SCENES = [
    ("intro", ["intro", "인트로", "오프닝"]),
    ("facility", ["facility", "facilities", "파실리티", "시설", "pool", "gym"]),
    ("small", ["small", "작은", "second", "bed2"]),
    ("big", ["big", "large", "master", "큰", "main", "bed1"]),
    ("closing", ["closing", "outro", "클로징", "엔딩"]),
]
EXTS = {".mp4", ".mov", ".m4v", ".mkv"}


def classify(name):
    low = name.lower()
    for idx, (_, keys) in enumerate(SCENES):
        if any(k in low for k in keys):
            return idx
    return None


def has_audio(path):
    out = subprocess.run(
        ["ffprobe", "-v", "error", "-select_streams", "a", "-show_entries",
         "stream=index", "-of", "csv=p=0", str(path)],
        capture_output=True, text=True)
    return bool(out.stdout.strip())


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("folder")
    ap.add_argument("-o", "--output", default="tour.mp4")
    ap.add_argument("--size", default="1920x1080")
    ap.add_argument("--fps", type=int, default=30)
    a = ap.parse_args()

    w, h = a.size.split("x")
    files = sorted(p for p in Path(a.folder).iterdir() if p.suffix.lower() in EXTS)
    ordered, unknown = [], []
    for p in files:
        idx = classify(p.name)
        (unknown if idx is None else ordered).append((idx, p) if idx is not None else p)
    if unknown:
        sys.exit("장면을 알 수 없는 파일: " + ", ".join(p.name for p in unknown)
                 + "\n파일명에 intro/facility/small/big/closing 중 하나를 넣어 주세요.")
    ordered.sort(key=lambda t: (t[0], t[1].name))
    clips = [p for _, p in ordered]
    if not clips:
        sys.exit("영상 파일이 없습니다.")

    print("조립 순서:")
    for i, p in enumerate(clips, 1):
        print(f"  {i}. {p.name}  [{SCENES[classify(p.name)][0]}]")
    missing = [n for n, (s, _) in enumerate(SCENES) if not any(classify(p.name) == n for p in clips)]
    for n in missing:
        print(f"  (참고) '{SCENES[n][0]}' 장면 클립이 없습니다.")

    cmd = ["ffmpeg", "-y", "-loglevel", "error", "-stats"]
    parts, labels = [], []
    for i, p in enumerate(clips):
        silent = not has_audio(p)
        cmd += ["-i", str(p)]
        if silent:
            cmd += ["-f", "lavfi", "-i", "anullsrc=r=48000:cl=stereo"]
    inp = 0
    for i, p in enumerate(clips):
        v = inp
        silent = not has_audio(p)
        a_idx = inp + 1 if silent else inp
        inp += 2 if silent else 1
        parts.append(f"[{v}:v]scale={w}:{h}:force_original_aspect_ratio=decrease,"
                     f"pad={w}:{h}:(ow-iw)/2:(oh-ih)/2,setsar=1,fps={a.fps},format=yuv420p[v{i}]")
        if silent:
            dur = subprocess.run(
                ["ffprobe", "-v", "error", "-show_entries", "format=duration",
                 "-of", "csv=p=0", str(p)], capture_output=True, text=True).stdout.strip()
            parts.append(f"[{a_idx}:a]atrim=0:{dur},asetpts=PTS-STARTPTS[a{i}]")
        else:
            parts.append(f"[{a_idx}:a]aresample=48000,aformat=channel_layouts=stereo,asetpts=PTS-STARTPTS[a{i}]")
        labels.append(f"[v{i}][a{i}]")
    parts.append("".join(labels) + f"concat=n={len(clips)}:v=1:a=1[v][a]")
    cmd += ["-filter_complex", ";".join(parts), "-map", "[v]", "-map", "[a]",
            "-c:v", "libx264", "-crf", "20", "-preset", "medium",
            "-c:a", "aac", "-b:a", "192k", "-movflags", "+faststart", a.output]
    subprocess.run(cmd, check=True)
    print("완료:", a.output)


if __name__ == "__main__":
    main()

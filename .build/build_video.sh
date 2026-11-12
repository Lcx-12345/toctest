#!/bin/bash
set -e
A=/workspace/ai-trends-2026/assets
V=/workspace/ai-trends-2026/video
FB=/usr/share/fonts/opentype/noto-cjk-otf/NotoSansCJKsc-Bold.otf
FR=/usr/share/fonts/opentype/noto-cjk-otf/NotoSansCJKsc-Regular.otf
[ -f "$FR" ] || FR=$FB

COMMON="scale=1600:900:force_original_aspect_ratio=increase,crop=1600:900"
ZP="zoompan=z='min(zoom+0.0007,1.10)':x='iw/2-(iw/zoom/2)':y='ih/2-(ih/zoom/2)':d=1:s=1280x720:fps=25"

# ---------- 1 片头 ----------
ffmpeg -y -loop 1 -framerate 25 -t 5.0 -i "$A/cover.jpg" -vf \
"$COMMON,gblur=sigma=22,eq=brightness=-0.20:saturation=0.85,$ZP,\
drawtext=fontfile=$FB:text='2026 AI 应用趋势':fontcolor=white:fontsize=62:x=(w-text_w)/2:y=(h-text_h)/2-46:shadowcolor=black@0.45:shadowx=3:shadowy=3,\
drawtext=fontfile=$FB:text='从「能用」到「常用」':fontcolor=0x4FE3D2:fontsize=34:x=(w-text_w)/2:y=(h-text_h)/2+40,\
drawtext=fontfile=$FR:text='趋势研究 · 2026.11':fontcolor=white@0.55:fontsize=18:x=(w-text_w)/2:y=h-86" \
-c:v libx264 -preset medium -crf 19 -pix_fmt yuv420p -an "$V/s1.mp4" -loglevel error

# ---------- 2 智能体 ----------
ffmpeg -y -loop 1 -framerate 25 -t 4.0 -i "$A/agent.jpg" -vf \
"$COMMON,$ZP,\
drawbox=x=0:y=0:w=iw:h=ih:color=black@0.22:t=fill,\
drawtext=fontfile=$FB:text='01 智能体工作流':fontcolor=white:fontsize=48:x=104:y=h-190:shadowcolor=black@0.5:shadowx=2:shadowy=2,\
drawtext=fontfile=$FR:text='从问答到执行  调用工具  对结果负责':fontcolor=0xCFE9E6:fontsize=22:x=106:y=h-118,\
drawbox=x=80:y=h-202:w=6:h=64:color=0x3EE0D0@0.95:t=fill" \
-c:v libx264 -preset medium -crf 19 -pix_fmt yuv420p -an "$V/s2.mp4" -loglevel error

# ---------- 3 AI 工程 ----------
ffmpeg -y -loop 1 -framerate 25 -t 4.0 -i "$A/code.jpg" -vf \
"$COMMON,$ZP,\
drawbox=x=0:y=0:w=iw:h=ih:color=black@0.22:t=fill,\
drawtext=fontfile=$FB:text='02 AI 原生软件工程':fontcolor=white:fontsize=48:x=104:y=h-190:shadowcolor=black@0.5:shadowx=2:shadowy=2,\
drawtext=fontfile=$FR:text='需求拆解  测试生成  代码评审':fontcolor=0xCFE9E6:fontsize=22:x=106:y=h-118,\
drawbox=x=80:y=h-202:w=6:h=64:color=0x8B7BFF@0.95:t=fill" \
-c:v libx264 -preset medium -crf 19 -pix_fmt yuv420p -an "$V/s3.mp4" -loglevel error

# ---------- 4 端侧 ----------
ffmpeg -y -loop 1 -framerate 25 -t 4.0 -i "$A/edge.jpg" -vf \
"$COMMON,$ZP,\
drawbox=x=0:y=0:w=iw:h=ih:color=black@0.18:t=fill,\
drawtext=fontfile=$FB:text='03 端侧小模型与 AI 硬件':fontcolor=white:fontsize=46:x=104:y=h-190:shadowcolor=black@0.5:shadowx=2:shadowy=2,\
drawtext=fontfile=$FR:text='隐私  时延  离线可用':fontcolor=0xCFE9E6:fontsize=22:x=106:y=h-118,\
drawbox=x=80:y=h-202:w=6:h=64:color=0x3EE0D0@0.95:t=fill" \
-c:v libx264 -preset medium -crf 19 -pix_fmt yuv420p -an "$V/s4.mp4" -loglevel error

# ---------- 5 片尾 ----------
ffmpeg -y -loop 1 -framerate 25 -t 5.0 -i "$A/cover.jpg" -vf \
"$COMMON,gblur=sigma=26,eq=brightness=-0.26:saturation=0.7,$ZP,\
drawtext=fontfile=$FB:text='把 AI 变成日常的能力':fontcolor=white:fontsize=54:x=(w-text_w)/2:y=(h-text_h)/2-20:shadowcolor=black@0.45:shadowx=3:shadowy=3,\
drawtext=fontfile=$FR:text='2026 AI 应用趋势报告':fontcolor=0x4FE3D2:fontsize=24:x=(w-text_w)/2:y=(h-text_h)/2+58" \
-c:v libx264 -preset medium -crf 19 -pix_fmt yuv420p -an "$V/s5.mp4" -loglevel error

# ---------- 合成 ----------
ffmpeg -y -i "$V/s1.mp4" -i "$V/s2.mp4" -i "$V/s3.mp4" -i "$V/s4.mp4" -i "$V/s5.mp4" -filter_complex \
"[0:v][1:v]xfade=transition=fade:duration=0.6:offset=4.4[x1];\
[x1][2:v]xfade=transition=fade:duration=0.6:offset=7.8[x2];\
[x2][3:v]xfade=transition=fade:duration=0.6:offset=11.2[x3];\
[x3][4:v]xfade=transition=fade:duration=0.6:offset=14.6,format=yuv420p[v]" \
-map "[v]" -c:v libx264 -preset slow -crf 20 -movflags +faststart -r 25 \
"/workspace/ai-trends-2026/2026 AI 应用趋势.mp4" -loglevel error

rm -f "$V"/s*.mp4
echo "done"
ffprobe -v error -show_entries format=duration,size -of default=nw=1 "/workspace/ai-trends-2026/2026 AI 应用趋势.mp4"

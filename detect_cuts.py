import os
import sys
import argparse
import torch
import cv2
import pandas as pd
import subprocess
import xlsxwriter
from transnetv2_pytorch import TransNetV2

def format_timestamp(seconds):
    ms = int((seconds - int(seconds)) * 1000)
    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    seconds = int(seconds % 60)
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}.{ms:03d}"

def get_video_info(video_path):
    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        return None, None
    fps = cap.get(cv2.CAP_PROP_FPS)
    total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
    cap.release()
    return fps, total_frames

def create_proxy(source_path, proxy_path):
    print(f"--- Creating analysis proxy (720p) ---")
    # Using h264_nvenc for speed if possible, fallback to libx264
    cmd = ['ffmpeg', '-y', '-hwaccel', 'auto', '-i', source_path, 
           '-vf', 'scale=1280:-1', '-c:v', 'libx264', '-preset', 'ultrafast', 
           '-crf', '28', '-an', proxy_path]
    subprocess.run(cmd, capture_output=True)

def generate_html(output_dir, video_name, shot_data, num_shots, total_frames):
    html_path = os.path.join(output_dir, "Shot_Sheet.html")
    html_content = f"""
<!DOCTYPE html>
<html>
<head>
    <title>{video_name} - Shot Sheet</title>
    <style>
        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #121212; color: #e0e0e0; padding: 30px; }}
        table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
        th {{ background: #1f1f1f; padding: 15px; text-align: left; border-bottom: 2px solid #444; }}
        td {{ padding: 15px; border-bottom: 1px solid #2a2a2a; vertical-align: middle; }}
        img {{ width: 320px; border-radius: 6px; border: 1px solid #333; }}
        .shot-code {{ font-family: monospace; font-weight: bold; font-size: 1.2em; color: #ffcc00; }}
        .timestamp {{ font-family: monospace; color: #00ffcc; }}
        textarea {{ width: 100%; height: 60px; background: #222; color: white; border: 1px solid #444; border-radius: 4px; padding: 5px; }}
    </style>
</head>
<body>
    <h1>Visual Shot Sheet: {video_name}</h1>
    <p>Total Shots: {num_shots} | Total Frames: {total_frames}</p>
    <table>
        <thead><tr><th>Shot Code</th><th>Thumbnail</th><th>Timestamp</th><th>Duration</th><th>Has VFX</th><th>Description</th></tr></thead>
        <tbody>
"""
    for shot in shot_data:
        html_content += f"""
        <tr>
            <td class="shot-code">{shot['Shot Code']}</td>
            <td><a href="thumbnails/{shot['Thumbnail']}" target="_blank"><img src="thumbnails/{shot['Thumbnail']}"></a></td>
            <td class="timestamp">{shot['Start (HH:MM:SS.ms)']}</td>
            <td>{shot['Frame Count']} frames</td>
            <td style="text-align: center;"><input type="checkbox" style="transform: scale(1.5);"></td>
            <td><textarea placeholder="Add description..."></textarea></td>
        </tr>"""
    html_content += "</tbody></table></body></html>"
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html_content)

def generate_xlsx(output_dir, shot_data):
    excel_path = os.path.join(output_dir, "Shot_List_Visual.xlsx")
    thumb_dir = os.path.join(output_dir, "thumbnails")
    workbook = xlsxwriter.Workbook(excel_path)
    worksheet = workbook.add_worksheet("Shot List")
    
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
    cell_fmt = workbook.add_format({'valign': 'vcenter', 'border': 1})
    
    headers = ["Shot Code", "Thumbnail", "Start (Seconds)", "Start (HH:MM:SS.ms)", "Frame Count", "Has VFX", "VFX Description"]
    for col, h in enumerate(headers): worksheet.write(0, col, h, header_fmt)
    
    worksheet.set_column('A:A', 15)
    worksheet.set_column('B:B', 30)
    worksheet.set_column('C:E', 18)
    worksheet.set_column('F:F', 12)
    worksheet.set_column('G:G', 40)
    
    for i, shot in enumerate(shot_data):
        row = i + 1
        worksheet.set_row(row, 80)
        worksheet.write(row, 0, shot['Shot Code'], cell_fmt)
        worksheet.write(row, 2, shot['Start (Seconds)'], cell_fmt)
        worksheet.write(row, 3, shot['Start (HH:MM:SS.ms)'], cell_fmt)
        worksheet.write(row, 4, shot['Frame Count'], cell_fmt)
        worksheet.write(row, 5, "", cell_fmt)
        worksheet.write(row, 6, "", cell_fmt)
        worksheet.data_validation(row, 5, row, 5, {'validate': 'list', 'source': ['Yes', 'No']})
        
        img_path = os.path.join(thumb_dir, shot['Thumbnail'])
        if os.path.exists(img_path):
            worksheet.insert_image(row, 1, img_path, {'x_scale': 0.1, 'y_scale': 0.1, 'x_offset': 5, 'y_offset': 5, 'object_position': 1})
            
    workbook.close()

def main():
    parser = argparse.ArgumentParser(description="AI Scene Detection and Thumbnail Extraction")
    parser.add_argument("video", help="Path to source video file")
    parser.add_argument("--output", help="Output root directory", default=".")
    args = parser.parse_args()

    video_path = os.path.abspath(args.video)
    video_name = os.path.splitext(os.path.basename(video_path))[0]
    proj_dir = os.path.join(os.path.abspath(args.output), f"{video_name}_scenedetect")
    thumb_dir = os.path.join(proj_dir, "thumbnails")
    os.makedirs(thumb_dir, exist_ok=True)

    print(f"--- Processing: {video_name} ---")
    device = "cuda" if torch.cuda.is_available() else "cpu"
    print(f"--- Using Device: {device} ---")
    
    model = TransNetV2(device=device)
    fps, total_frames = get_video_info(video_path)
    if not fps:
        print("Error: Could not read video file.")
        return

    proxy_path = os.path.join(proj_dir, "temp_proxy.mp4")
    create_proxy(video_path, proxy_path)
    
    print("--- Analyzing scenes ---")
    scenes = model.detect_scenes(proxy_path)
    os.remove(proxy_path)
    
    num_shots = len(scenes)
    print(f"--- Detected {num_shots} shots ---")
    
    shot_data = []
    for i in range(num_shots):
        shot_num = i + 1
        shot_code = f"sh_{shot_num * 10:04d}"
        
        scene_info = scenes[i]
        start_frame = scene_info.get('start', scene_info[0]) if isinstance(scene_info, (dict, list, tuple)) else scene_info
        
        if i < num_shots - 1:
            next_info = scenes[i+1]
            next_start = next_info.get('start', next_info[0]) if isinstance(next_info, (dict, list, tuple)) else next_info
            end_frame = next_start
        else:
            end_frame = total_frames
            
        ts_sec = float(start_frame) / fps
        thumb_name = f"{shot_code}.jpg"
        thumb_path = os.path.join(thumb_dir, thumb_name)
        
        # Extract thumbnail
        cmd = ['ffmpeg', '-y', '-loglevel', 'error', '-ss', str(ts_sec), '-i', video_path, 
               '-vframes', '1', '-vf', 'scale=1920:-1', '-q:v', '2', thumb_path]
        subprocess.run(cmd)
        
        shot_data.append({
            "Shot Code": shot_code, "Start (Seconds)": round(ts_sec, 3),
            "Start (HH:MM:SS.ms)": format_timestamp(ts_sec),
            "Frame Count": int(end_frame - start_frame), "Thumbnail": thumb_name
        })
        if shot_num % 50 == 0: print(f"Extracted {shot_num}/{num_shots} thumbnails...")

    pd.DataFrame(shot_data).to_csv(os.path.join(proj_dir, "shot_list.csv"), index=False)
    generate_html(proj_dir, video_name, shot_data, num_shots, total_frames)
    generate_xlsx(proj_dir, shot_data)
    print(f"--- Complete! Results in: {proj_dir} ---")

if __name__ == "__main__":
    main()

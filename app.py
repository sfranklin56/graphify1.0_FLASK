import os
import tempfile
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import cv2
from flask import Flask, request, send_file, render_template
from matplotlib.animation import FuncAnimation

app = Flask(__name__)

# Allowed file types
ALLOWED_EXTENSIONS = {'xls', 'xlsx'}

# Check if uploaded file is valid
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files.get("file")
        plot_type = request.form.get("plot_type")
        plot_title = request.form.get("plot_title")

        if not file or not allowed_file(file.filename):
            return "Invalid file format. Please upload an Excel file.", 400

        # Read Excel data
        df = pd.read_excel(file)
        
        # Generate a video based on the selected plot type
        temp_video = generate_video(df, plot_type, plot_title)

        return send_file(temp_video, mimetype="video/mp4", as_attachment=True, download_name="animation.mp4")

    return render_template("index.html")

def generate_video(df, plot_type, plot_title):
    fig, ax = plt.subplots()

    def update(frame):
        ax.clear()
        ax.set_title(plot_title)
        
        if plot_type == "line":
            ax.plot(df.iloc[:frame, 0], df.iloc[:frame, 1], marker="o")
        elif plot_type == "bar":
            ax.bar(df.iloc[:frame, 0], df.iloc[:frame, 1])
        elif plot_type == "pie":
            ax.pie(df.iloc[frame, 1:], labels=df.columns[1:], autopct="%1.1f%%")
    
    frames = min(30, len(df))
    ani = FuncAnimation(fig, update, frames=frames, repeat=False)

    # Save as MP4
    temp_dir = tempfile.mkdtemp()
    video_path = os.path.join(temp_dir, "output.mp4")
    ani.save(video_path, fps=5, extra_args=['-vcodec', 'libx264'])

    return video_path

if __name__ == "__main__":
    app.run(debug=True)

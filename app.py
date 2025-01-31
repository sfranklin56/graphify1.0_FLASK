import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.animation import FuncAnimation
from moviepy.editor import VideoFileClip  # Import moviepy for exporting the animation
from flask import Flask, render_template, request
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['ALLOWED_EXTENSIONS'] = {'xlsx', 'xls'}
app.config['STATIC_FOLDER'] = 'static'

SPEED_OPTIONS = {
    'slowest': 1000,
    'slower': 750,
    'average': 500,
    'faster': 250,
    'fastest': 100
}

SIZE_OPTIONS = {
    'YouTube': (16, 9),
    'Instagram': (1, 1),
    'YouTube Shorts': (9, 16),
    'Mobile': (9, 16),
    'PC': (16, 9)
}

QUALITY_OPTIONS = {
    '480p': 30,
    '720p': 60,
    '1080p': 120,
    '4k': 240
}

if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            return 'No file part', 400

        file = request.files['file']
        if file.filename == '':
            return 'No selected file', 400

        # Save the uploaded file
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file.filename))
        file.save(file_path)

        plot_type = request.form.get('plot_type', 'line')
        plot_title = request.form.get('plot_title', 'Plot Title')
        speed_option = request.form.get('speed', 'average')
        invert_colors = request.form.get('invert_colors') is not None
        size_option = request.form.get('size', 'YouTube')
        video_quality = request.form.get('video_quality', '720p')

        # Animation speed and size
        animation_speed = SPEED_OPTIONS.get(speed_option, 500)
        width, height = SIZE_OPTIONS.get(size_option, (16, 9))

        try:
            df = pd.read_excel(file_path)
        except Exception as e:
            return f"Error reading the Excel file: {e}", 400

        date_col = 'Date'
        if date_col not in df.columns:
            return "Date column is missing in the data.", 400

        df[date_col] = pd.to_datetime(df[date_col], format='%Y', errors='coerce')
        df = df.dropna(subset=[date_col])  # Drop rows with invalid date
        dates = df[date_col]
        dates_numeric = dates.astype(np.int64)
        df = df.fillna(0)  # Fill NaN with zero

        data_columns = df.columns[df.columns != date_col]

        # Interpolating data
        interpolated_data = {}
        for col in data_columns:
            y_data = pd.to_numeric(df[col], errors='coerce').fillna(0)
            new_dates = pd.date_range(start=dates.min(), end=dates.max(), freq='h')
            new_dates_numeric = new_dates.astype(np.int64)
            new_y_data = np.interp(new_dates_numeric, dates_numeric, y_data)
            interpolated_data[col] = new_y_data

        fig, ax = plt.subplots(figsize=(width, height))

        if invert_colors:
            ax.set_facecolor('#333333')
            plt.rcParams['axes.facecolor'] = '#333333'
            plt.rcParams['text.color'] = 'white'
            plt.rcParams['axes.labelcolor'] = 'white'
            plt.rcParams['xtick.color'] = 'white'
            plt.rcParams['ytick.color'] = 'white'
        else:
            ax.set_facecolor('#ffffff')

        plt.style.use("ggplot")
        new_dates = pd.date_range(start=dates.min(), end=dates.max(), freq='h')
        num_points = len(new_dates)

        # Animation function
        def animate(i):
            ax.clear()
            if invert_colors:
                ax.set_facecolor('#333333')
            else:
                ax.set_facecolor('#ffffff')

            if plot_type == 'line':
                for col in data_columns:
                    ax.plot(new_dates[:i+1], interpolated_data[col][:i+1], label=col, linewidth=4, linestyle='-', marker='o', markersize=4)
                    ax.text(new_dates[i], interpolated_data[col][i], col, color='black' if not invert_colors else 'white', fontsize=10, ha='left', va='bottom')
            elif plot_type == 'bar':
                current_data = {col: interpolated_data[col][i] for col in data_columns}
                sorted_data = sorted(current_data.items(), key=lambda x: x[1], reverse=True)
                categories, values = zip(*sorted_data)
                ax.barh(categories, values, color='#1f77b4')
                ax.set_xlim(0, max(values) * 1.1)
            elif plot_type == 'pie':
                current_data = {col: interpolated_data[col][i] for col in data_columns}
                sizes = list(current_data.values())
                labels = list(current_data.keys())
                ax.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=140)

            ax.legend()
            plt.tight_layout()

        ani = FuncAnimation(fig, animate, frames=num_points, interval=animation_speed, repeat=False)

        # Save animation to a temporary file
        video_file_path = os.path.join(app.config['STATIC_FOLDER'], 'output_video.mp4')
        ani.save(video_file_path, writer='ffmpeg', dpi=QUALITY_OPTIONS.get(video_quality, 60))

        return render_template('index.html', video_file='output_video.mp4')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)

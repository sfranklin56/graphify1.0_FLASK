import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.animation import FuncAnimation
from flask import Flask, render_template, request, send_from_directory
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['ALLOWED_EXTENSIONS'] = {'xlsx', 'xls'}
app.config['STATIC_FOLDER'] = 'static'

# Allowed file extensions
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

# Predefined speed options (milliseconds per frame)
SPEED_OPTIONS = {
    'slowest': 1000,
    'slower': 750,
    'average': 500,
    'faster': 250,
    'fastest': 100
}

# Predefined sizes for different platforms
SIZE_OPTIONS = {
    'YouTube': (16, 9),
    'Instagram': (1, 1),
    'YouTube Shorts': (9, 16),
    'Mobile': (9, 16),
    'PC': (16, 9)
}

# Predefined video quality settings
QUALITY_OPTIONS = {
    '480p': 30,
    '720p': 60,
    '1080p': 120,
    '4k': 240
}

# Ensure the upload folder exists
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

# Route for the homepage
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

        # Get form data
        plot_type = request.form.get('plot_type', 'line')
        plot_title = request.form.get('plot_title', 'Plot Title')
        speed_option = request.form.get('speed', 'average')  # Default speed is 'average'
        invert_colors = request.form.get('invert_colors') is not None  # Check if invert colors is selected
        size_option = request.form.get('size', 'YouTube')
        video_quality = request.form.get('video_quality', '720p')

        # Set animation speed based on selection
        animation_speed = SPEED_OPTIONS.get(speed_option, 500)

        # Get dimensions based on selected size
        width, height = SIZE_OPTIONS.get(size_option, (16, 9))

        x_label = request.form.get('x_label', 'X-axis') if plot_type in ['line', 'bar'] else ''
        y_label = request.form.get('y_label', 'Y-axis') if plot_type in ['line', 'bar'] else ''

        try:
            df = pd.read_excel(file_path)
        except Exception as e:
            return f"Error reading the Excel file: {e}", 400

        date_col = 'Date'
        if date_col not in df.columns:
            return "Date column is missing in the data.", 400

        df[date_col] = pd.to_datetime(df[date_col], format='%Y', errors='coerce')
        df = df.dropna(subset=[date_col])  # Drop rows where the Date column is invalid
        dates = df[date_col]
        dates_numeric = dates.astype(np.int64)
        df = df.fillna(0)  # Fill NaN values with 0

        data_columns = df.columns[df.columns != date_col]

        # Prepare interpolated data
        interpolated_data = {}
        for col in data_columns:
            y_data = pd.to_numeric(df[col], errors='coerce').fillna(0)
            new_dates = pd.date_range(start=dates.min(), end=dates.max(), freq='h')
            new_dates_numeric = new_dates.astype(np.int64)
            new_y_data = np.interp(new_dates_numeric, dates_numeric, y_data)
            interpolated_data[col] = new_y_data

        fig, ax = plt.subplots(figsize=(width, height))

        # Dark theme settings for inverted colors
        if invert_colors:
            ax.set_facecolor('#333333')  # Dark background for graph
            plt.rcParams['axes.facecolor'] = '#333333'
            plt.rcParams['text.color'] = 'white'
            plt.rcParams['axes.labelcolor'] = 'white'
            plt.rcParams['xtick.color'] = 'white'
            plt.rcParams['ytick.color'] = 'white'
        else:
            ax.set_facecolor('#ffffff')  # Light background for graph

        plt.style.use("ggplot")
        new_dates = pd.date_range(start=dates.min(), end=dates.max(), freq='h')
        num_points = len(new_dates)

        # Function to animate the plot
        def animate(i):
            ax.clear()
            if invert_colors:
                ax.set_facecolor('#333333')  # Dark background for graph
            else:
                ax.set_facecolor('#ffffff')  # Light background for graph

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
                ax.set_xlabel(x_label, color='black' if not invert_colors else 'white')
                ax.set_ylabel(y_label, color='black' if not invert_colors else 'white')
                ax.invert_yaxis()
            elif plot_type == 'pie':
                current_data = {col: interpolated_data[col][i] for col in data_columns}
                sizes = list(current_data.values())
                labels = list(current_data.keys())
                ax.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=140)

            ax.legend()
            plt.tight_layout()

        # Set up the animation
        ani = FuncAnimation(fig, animate, frames=num_points, interval=animation_speed, repeat=False)

        # Save the animation as an MP4 file
        video_file_path = os.path.join(app.config['STATIC_FOLDER'], 'output_video.mp4')
        ani.save(video_file_path, writer='ffmpeg', dpi=QUALITY_OPTIONS.get(video_quality, 60))

        return render_template('index.html', video_file='output_video.mp4')

# Run the app on the right port for deployment
if __name__ == '__main__':
    # Use the PORT from the environment variable or default to 5000 for local testing
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)

# GestureDeck

GestureDeck enables users to control PowerPoint presentations seamlessly using hand gestures via a webcam. This project combines the power of OpenCV and Mediapipe for real-time hand gesture recognition and system interaction.

## Features
- **Slide Navigation**: Swipe gestures to move between slides.
- **Pointer Control**: Use your hand to move the mouse pointer.
- **Zoom In/Out**: Perform pinch gestures to zoom in and out during the presentation.
- **Marker Tool**: Draw annotations on the presentation when enabled.

## Setup

### Prerequisites
Ensure you have the following installed:
- Python 3.7 or higher
- A webcam for gesture detection
- Microsoft PowerPoint (2016 or newer)

### Installation
1. Clone the repository:
   ```bash
   git clone https://github.com/your-username/gesturedeck.git
   cd gesturedeck
   ```

2. Install the required Python packages:
   ```bash
   pip install opencv-python mediapipe pyautogui numpy pywin32
   ```

3. Update the PowerPoint file path in the scripts (`CustomModules.py` and `MediapipeCode.py`):
   Replace `C:\Users\ADITYA\Desktop\hand-gesture-control\presentation.pptx` with the path to your PowerPoint file.

## Usage

1. Run the script:
   ```bash
   python CustomModules.py
   ```
   or
   ```bash
   python MediapipeCode.py
   ```

2. Use the following gestures to control your presentation:
   - **Swipe Left/Right**: Navigate between slides.
   - **Index Finger Point**: Control the pointer.
   - **Pinch Gesture**: Zoom in/out within the presentation.
   - **Thumb Direction**: Move slides forward or backward.

3. Press `q` to exit the application.

## Troubleshooting
- Ensure your webcam is functional.
- Verify the PowerPoint file path is correct.
- Check the terminal for error messages and resolve any issues accordingly.

## Contributing
Feel free to fork this repository, make changes, and submit a pull request.

## License
This project is licensed under the MIT License.


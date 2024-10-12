# SolidWorks-Rich-Presence

## Description
A program that utilizes Python libraries to add Discord rich presence to SolidWorks, so your mates can see what you're currently working on.

## Installation

### Prerequisites
- Python 3 
- Pip

### Steps
1. Clone the repository:
    ```bash
    git clone https://github.com/chxdr/SolidWorks-Rich-Presence
    cd SolidWorks-Rich-Presence
    ```

2. Install the required libraries:
    ```bash
    pip install psutil pygetwindow pypresence configparser pillow pywin32
    ```

3. Run the program:
    ```bash
    python main.py
    ```

## How it Works
- **psutil**: Tracks the running SolidWorks processes.
- **pygetwindow**: Captures the current window title from SolidWorks to show what you're working on.
- **pypresence**: Communicates with Discord's [Rich Presence API](https://discord.com/developers/docs/rich-presence/how-to) to update your status.
- **configparser**: Handles configuration settings for easier customization.
- **pillow**: Used for adding custom image assets to the rich presence.
- **pywin32**: Interacts with Windows to handle SolidWorks-specific window operations.

## Notes
I did originally create this program a few years back before I actually learnt to program using ChatGPT, so at some point I might rewrite this entire program myself.

And currently drawing files are not supported, not sure why but I'm not fixing it.

# Swayum

Swayum is an open-source Python script which provides a user-friendly graphical interface for capturing and executing mouse and keyboard actions, making it convenient for tasks that involve repetitive operations.It leverages the potential of pyautogui and makes it usable in a GUI form. SWAYAM (meaning 'Self' in Sanskrit) is also used synonymously with the word automatic.  

# Features:

1. Capture mouse positions and set actions accordingly to automate repeated mouse movements
2. Perform actions like single click, right click, and text input at the captured positions
3. Save and load captured positions for future
4. Preview captured positions on a graphical canvas
5. Intuitive interface with start/stop capturing buttons and action selection dialog
   
<img width="960" alt="capture" src="https://github.com/computerauditor/Swayum/assets/117805200/92eae0d5-5278-4a4c-b502-5cf16852c6b7">
**Basic User Interface**

# Instructions/How to use

- Python and pyautogui are default dependencies
- Use the below command to install pyautogui
```
pip install pyautogui
```
- Run this script and use "Ctrl+I" as hotkey to capture the mouse movement by clicking the "start recording" button
- You can stop recording once done and by using "Execute" button execute the action
- Also we can save or load the mouse movements via a simple .txt file
- Use "Preview Position" button to show a preview of mouse movements captured. 

# PREVIEW

https://github.com/computerauditor/Swayum/assets/117805200/f524429e-58dc-43c1-b13b-fb39fadbc18b

*Swayum [Basic use to perform some simmple clicks and typing text]*

![Preview](https://github.com/computerauditor/Swayum/assets/117805200/343edb51-a82f-4d1c-8150-9603bcb886fd)
*User can preview their mouse movement that they captured using the "Position Preview" graph*


![load and save](https://github.com/computerauditor/Swayum/assets/117805200/e8bf4372-9c90-4d07-90be-13406e4db799)
*User can save or load their mouse movements as a simple .txt file for future reference*

# Credits
Team pyautogui

# Note
Don't create issue for basic stuff like how to install python or related library. PR are always welcome!!!

# Upcomming features
1. To fix some bugs
2. To add more options such as menu for using special keys such as enter,shift,ctrl,alt
3. Enabling macro recording , so that user can start recording their action with mouse and keyboard and the script automatically captures them , thus enable user to reperform it alltogether with a single click for future reference.
4. Feel free to raise an idea or request to improve this script.

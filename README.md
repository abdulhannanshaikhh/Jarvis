# Jarvis
The virtual assistant is a voice-controlled virtual assistant, "Jarvis," that uses the speech_recognition package for voice input and the win32com.client module for text-to-speech output. It can open websites, display the time, and execute system actions via os and subprocess modules.

Modules Used and Library Requirements:-

1. `speech_recognition`:
   - This module provides functionalities to recognize speech from various sources, including microphones and audio files.
   - In the code, it is used to capture voice input from the user through the `Microphone` class.
   - The `recognize_google` method of the `Recognizer` class is used to convert the captured audio into text using Google's speech recognition API.

2. `win32com.client`:
   - This module allows Python to interact with COM (Component Object Model) objects on Windows.
   - It provides access to various Windows components, including the Windows Speech API (SAPI).
   - In the code, it is used to create a text-to-speech engine using the `Dispatch` function to enable the virtual assistant to respond audibly to the user's commands.

3. `os`:
   - This is a built-in module that provides a portable way of using operating system-dependent functionality.
   - In the code, it is used to interact with the operating system and open files and applications using the `startfile()` method.

4. `subprocess`:
   - This module allows the creation of additional processes and communication with them.
   - In the code, it is used to run system commands or launch applications using the `run()` function.

5. `webbrowser`:
   - This module provides a high-level interface to work with web browsers.
   - In the code, it is used to open URLs in the default web browser using the `open()` function.

6. `datetime`:
   - This is a built-in module that provides classes for manipulating dates and times.
   - In the code, it is used to get the current time using `datetime.now()` and then format it into hours and minutes.

7. `random`:
   - This is a built-in module that provides functions to generate random numbers and sequences.
   - In the code, it is imported but not actively used.

8. `numpy`:
   - This is a powerful third-party package for numerical computing in Python.
   - However, in the provided code, it is imported but not actively used.

9. `winshell`:
   - This module is not imported in the code, but it is used for handling interactions with the Windows Recycle Bin.

In summary, these modules and packages are utilized to enable voice recognition, text-to-speech functionality, system interactions, and web browsing capabilities in the virtual assistant application.

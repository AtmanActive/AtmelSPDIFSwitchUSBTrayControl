# Atmel S/PDIF Switch USB Tray Control

Windows system tray utility to control an Atmel based S/PDIF Switch created by Beni_Skate on Tindie:
https://www.tindie.com/products/beni_skate/automatic-spdif-opticalrca-audio-switch/

This S/PDIF switch has intelligent automatic switching by sensing inputs and their activity. But that is for the stuffed-behind-the-TV use case.
As I am not using it like that, but for my digital studio, where several computers and their digital outputs are connected to several digital devices (speakers, headphones...), I needed a utility to explicitly control the inputs on demand. So I built one.

This utility does not poll the device in any way. It is not aware of what is going on inside the device. It just sends serial COM commands blindly and that's it. Still, for choosing the desired input on demand - good enough. No drivers are required. It is enough to connect the device via USB cable and that's it.

I built the program using AutoIt3 and there is a pre-built x64 binary in the dist folder. The program is built with Portable Paradigm in mind, meaning, there is no install or uninstall or writing to registry or hidden files or any of that stuff. You can run the program from wherever you like and it will work as expected.

To configure the program, just edit the included ini file. It should be self-explanatory. You need to set the COM port name and channel names. Optionally, you can set custom channel icons to override the built-in ones. Also, there is on-run and on-exit section to switch to a desired channel on startup and exit respectively.

Enjoy.

AtmanActive 2020, 2021.

https://github.com/AtmanActive/AtmelSPDIFSwitchUSBTrayControl

v1.1: added faster detection timeout to make it snappier when switching inputs

v1.2: changed COM port open/close logic to on-demand to make it snappier when switching inputs

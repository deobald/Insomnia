# Insomnia

A graphical MUD I wrote in 1996

It might be possible to get this running on Windows, but I'm not confident.
Doing so might require force-feeding Windows 10/11 the old VB5 runtime and OCX files?
The following instructions are for Linux/Ubuntu.
Unfortunately, MacOS 10.15+ no longer supports 32-bit programs, which means wine can't run Win32 apps on MacOS 11.x.

## Required: Wine

```
sudo apt install wine
```

## Required: Winetricks

TODO: I installed a pile of things by hand, in addition to this... need to repeat.

```
wget http://winetricks.org/winetricks
chmod +x winetricks
sh winetricks corefonts vcrun6 vb5run native_oleaut32 vcrun2010 richtx32
```

## Required: OCX

* `MSWINSCK.OCX` - https://www.ocxdump.com/download-ocx-files_new.php/ocxfiles/M/MSWINSCK.OCX/6.00.81694/download.html
* `MCI32.OCX` - https://www.ocxdump.com/download-ocx-files_new.php/ocxfiles/M/Mci32.ocx/6.00.8418/download.html

```
# check if you have them already:
ls ~/.wine/drive_c/windows/system32/MSWINSCK* ~/.wine/drive_c/windows/system32/mci*

# copy downloaded files:
cp ~/Downloads/MSWINSCK.OCX ~/.wine/drive_c/windows/system32/MSWINSCK.OCX
cp ~/Downloads/Mci32.ocx ~/.wine/drive_c/windows/system32/Mci32.ocx
```

## Running

Clone this repo.

```
# run a server and grab the IP that shows up in the GUI
cd Insomnia/Mud/Server/Build004
WINEARCH="win32" wine Server04-02.exe

# connect the client to the server's IP
cd Insomnia/Mud/Client/Build009
WINEARCH="win32" wine Ins009-06.exe
```

# How to Unlock All Levels in MWCASexVR using UnityExplorer

Since you're starting fresh, follow these exact steps to install **BepInEx 6 for IL2CPP** and the latest **UnityExplorer** for Unity 6000. This will allow you to explore the game's internal structure and unlock locked content.

### Step 1: Clean Up Previous Attempts
1. Go to your game folder (`Married.Woman.Countdown.ASMR.Sex.VR\game`).
2. If you see a `BepInEx` folder, a `winhttp.dll` file, or a `doorstop_config.ini` file, **delete them**. We want a clean slate.

### Step 2: Download the Correct BepInEx
Since your game runs on **Unity 6000** and is compiled using **IL2CPP** (evident from your screenshot showing `GameAssembly.dll`), you *must* use BepInEx 6 for IL2CPP. BepInEx 5 will not work.

1. Go to the [BepInEx Bleeding Edge builds site](https://builds.bepinex.dev/projects/bepinex_be).
2. Look for the latest successful build (usually at the very top).
3. Click the **Artifacts** dropdown and download the file ending in `..._IL2CPP_x64.zip`.
   *(Example: `BepInEx_UnityIL2CPP_x64_9c2b17f_6.0.0-be.755.zip`)*

### Step 3: Install BepInEx
1. Open the `.zip` file you just downloaded.
2. Select everything inside the `.zip` (the `BepInEx` folder, `winhttp.dll`, and `doorstop_config.ini`).
3. Drag and drop them directly into your game folder, right next to `MWCASexVR.exe`.
4. **First Run:** Launch `MWCASexVR.exe` once. You should see a black console window pop up. Let the game fully load to the main menu.
5. Close the game.
   *(This first run allows BepInEx to generate necessary files inside the `BepInEx` folder, specifically the `Interop` assemblies for your game).*

### Step 4: Download UnityExplorer
Now we need the tool that will actually let you manipulate the game to unlock levels.

1. Go to the official [UnityExplorer Releases Page](https://github.com/sinai-dev/UnityExplorer/releases).
2. Under the latest release (usually v4.9.0 or newer), download:
   - `UnityExplorer.BepInEx.IL2CPP.CoreCLR.zip`
   *(Because your game uses BepInEx 6 with a .NET 6.0 runtime, you MUST use the CoreCLR version).*

### Step 5: Install UnityExplorer
1. Open the `UnityExplorer.BepInEx.IL2CPP.CoreCLR.zip` file.
2. Inside the `.zip`, you will see a `plugins` folder.
3. Open your game's `BepInEx` folder (`Married.Woman.Countdown.ASMR.Sex.VR\game\BepInEx`).
4. Drag and drop the `plugins` folder from the `.zip` into the `BepInEx` folder.
   *(This will merge the folders. Inside `BepInEx\plugins`, you should now have `UnityExplorer.BepInEx.IL2CPP.CoreCLR.dll` and `UniverseLib.IL2CPP.Interop.dll`)*.

### Step 6: Using UnityExplorer to Unlock Levels
1. Launch `MWCASexVR.exe`.
2. The BepInEx console should appear, and this time, it should eventually say `1 plugins to load` (referring to UnityExplorer).
3. When the game reaches the main menu, look at your desktop monitor (not just in VR). You should see the **UnityExplorer** UI panel on top of the game window.
   *(If your mouse is stuck, press the `F7` key to toggle the UnityExplorer interface and free your mouse).*
4. In the UnityExplorer window, click the **Object Search** tab at the top.
5. You need to look for the script or manager that controls game progress or saving.
   - Search for terms like `SaveData`, `GameManager`, `Unlock`, `Progress`, `Gallery`, or `Scenario`.
6. Once you find the relevant script/class, click on it. The **Inspector** tab will open on the right.
7. Look through the properties and fields in the Inspector. You are looking for boolean (True/False) values like `IsUnlocked`, `HasCleared`, `IsPlayed`, or integer arrays holding unlock states.
8. Change any `False` values to `True`, or change integers from `0` to `1` (or whatever the unlocked state is).
9. If there is an `UnlockAll()` or `ClearSave()` method (button) visible in the inspector, click **Invoke**.

*Note: Every game's internal structure is unique. You will have to use some intuition in the Object Search to find the exact class that holds the unlock states for MWCASexVR.*
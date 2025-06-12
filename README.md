# NOR_Analysis

This repository provides protocols for both manual and automated scoring of NOR (Novel Object Recognition) data. All code used for each scoring method can be found within respective manual or automated folders. Follow the steps below to set up your environment and run the analyses.

---

## 1. Python Environment Setup

Both scoring methods require a Python environment with specific packages installed. We recommend using miniconda (lightweight version of Anaconda)for environment management (https://www.anaconda.com/docs/getting-started/miniconda/main):

    1. Download and install miniconda 
    2. Open Terminal (macOS/Linux) or Anaconda Prompt (Windows)
    3. To create and activate new Python environment to be used for analysis: 
         conda create -n NORenv python=3.10
         conda activate NORenv
    4. Install packages: 
         conda install pandas pillow ffmpeg
         pip install openpyxl av==11

When required, this environment will be referred to as `NORenv` from this point forward and must be activated by entering following into Terminal/Anaconda Prompt: 
	 conda activate NORenv

---

## 2. Manual Scoring Method

### Prerequisites
* NORenv
* Microsoft Excel (with Office Scripts enabled)
* Chronotate-analyzed data in the form of .csv files. Refer to Chronotate GitHub wiki for further guidance (https://github.com/ShumanLab/Chronotate/wiki). 

### 2.1. Compile CSV Files & Format DI Data

Place all .csv files into one folder, then:

    1. Save `manual_DI_format.py` into directory containing .csv files.
    2. Open Terminal/Anaconda Prompt and enter the following:
       	 cd /Path/to/csv_folder #replace /Path with full path of .csv-containing directory
         conda activate NORenv		
	 python manual_DI_format.py -o manual_output.xlsx -p -v #replace `manual_output` with desired filename

NOTE: Excel sheet names in the resulting `manual_output.xlsx` will mirror .csv filenames. Since Excel limits sheet names to 31 characters, shorten filename if necessary.

---

### 2.2. Calculate Discrimination Index (DI)
  
    1. Open manual_output.xlsx in Excel
    2. Go to the Automate tab → Click New Script in Scripting Tools section
    3. Paste the contents of `manual_DI_OfficeScript.txt` into the editor
    4. Save and run the script

---

## 3. automated Scoring Method

The DLC model used in this protocol can be downloaded from this link: https://drive.google.com/drive/folders/1xnnoJec20t-z3MlZzoNtM3rGG4KqjDUb?usp=sharing

### Prerequisites
* NORenv
* MATLAB with required Toolboxes:
    - Video Processing
    - Curve-Fitting
    - Image Processing
    - Signal Processing
    - Statistics and Machine Learning
* Microsoft Excel (with VBA and Office Scripts enabled)
* BehaviorDEPOT-analyzed data in the form of folders containing required .mat data files. Refer to the BehaviorDEPOT GitHub Wiki for further guidance (https://github.com/DeNardoLab/BehaviorDEPOT/wiki).

### 3.1. Prepare Analysis Folders

After running BehaviorDEPOT:

    1. Consolidate all *_analyzed folders into one parent directory.
    2. Add source videos to corresponding *_analyzed folders
    3. Remove `_analyzed` from all subfolder names.

NOTE: `_Obj1` / `_Obj2`  should be appended onto subfolder names before running BehaviorDEPOT analysis in order to ensure smooth execution of below scripts. 
Excel sheet names in the resulting `compiled_behavior.xlsx` and `compiled_kinematics.xlsx` files will be copied from folder names. 
Since Excel limits sheet names to 31 characters, shorten folder names if necessary.

### 3.2. Extract Behavior & Kinematics Data from .mat data files

    1. Save `auto_extract_behavior_kinematics.m` MATLAB code in directory containing BehaviorDEPOT-analyzed data.
    2. Open MATLAB, locate MATLAB File Explorer and find analyzed directory 
    3. Right-click directory → Add to Path → Selected Folders and Subfolders
    4. Run the script by typing the following into MATLAB command window, replace `parentDir` with the full path of analyzed directory:  
	     auto_extract_behavior_kinematics(parentDir)
    5. Select analyzed directory when prompted, a completion message will be displayed upon successful execution of script.

This generates `compiled_behavior.xlsx` (containing frame numbers of all object exploration bouts for objects 1 & 2) and `compiled_kinematics.xlsx` (containing extracted total distance traveled, average velocities, and average accelerations from head and midback smoothed keypoint-tracking data) Excel files. If other or additional data points are to be extracted modify text under `T_kinematics = table()` within `extract_behavior_kinematics.m` code.

### 3.3. Format & Calculate DI

#### 3.3.1. Combine Obj1 and Obj2 Data 

    1. Save `auto_DI_format.py` in same directory as `compiled_behavior.xlsx`.
    2. Open Terminal/Anaconda Prompt and enter the following (replace '/Path/' sections with correct full paths of input/output files):
         cd /Path/to/compiled_behavior.xlsx
         conda activate NORenv
      	 python auto_DI_format.py 
	   --input "/Path/to/compiled_behavior.xlsx" \
  	   --output "/Path/to/output_file.xlsx"

#### 3.3.2. Calculate DI in Excel

    1. Open output_file.xlsx in Excel
    2. Go to the Automate tab → Click New Script in Scripting Tools section
    3. Clear editor of all contents, copy and paste the contents of `auto_OfficeScript_for_DI_calc.txt` into the editor
    4. Save and run the script, a completion message will appear if script was successfully executed.


### 3.4. Process Kinematic Data in Excel

    1. Open `compiled_kinematics.xlsx` in Excel.
    2. Go to Developer tab → Visual Basic under Code section. A separate Microsoft Visual Basic window will now be  displayed. 
    3. In this window, on left side locate the  VBAProject (compiled_kinematics.xisx) tree then select `ThisWorkbook`. Right click ThisWorkbook → Insert → Module. A text window on the right side will now be displayed.
    4. In the righthand window, copy and paste script within  `VBA_macro_for_kinematics.txt` file into window
    5. Close the left window then click the Run (▶️) button to execute macro. A message will confirm successful execution upon completion.

NOTE: Do not have any other Excel files opened during this step

---

## 4. Visualize automated-Scored Behavior
These steps overlay the smoothed keypoint tracking data (from nose, head, and tailpoint) and behavior label indicating object exploration was detected using automated-scoring method.
If automated-scored results differ from manual scoring this is a great place to start. 

### Prerequisites
* NORenv
* MATLAB with required Toolboxes:
    - Video Processing
    - Image Processing
* `Obj1` and `Obj2` BehaviorDEPOT-analyzed directories with source videos in all subfolders
  NOTE: The same folders from step 3 can be used for this step, seperate so that Obj1/Obj2 data is segregated

### 4.1. Create Labeled Videos

    1. Save `auto_processAllVideosObj1.m` MATLAB code within the `Obj1` directory.
    2. Add `Obj1` directory to MATLAB Path (as done in 3.2.2 - 3.2.3).
    3. Open `Obj1` directory in File Explorer and type following into MATLAB command window, replace `parentDir` with the full path of `Obj1` directory:
	       auto_processAllVideosObj1(parentDir)
    4. Repeat steps 4.1.1 - 4.1.3, this time saving `auto_processAllVideosObj2.m` MATLAB code in `Obj2` directory.

This step will add `Explore_Obj1`/ `Explore_Obj2` upper corner labels over video frames whenever object exploration was detected.

### 4.2. Overlay Cropped Behavior Label from Obj2 onto Obj1 Videos

#### 4.2.a. Crop `Explore_Obj2` Labels 
       	  
    1. Create a directory containing full-sized `Explore_Obj2`-labeled videos.
    2. Open Terminal/Anaconda Prompt and enter the following (replace '/Path/' with correct full path of video directory):
         cd /Path/to/Explore_Obj2_videoDirectory
         conda activate NORenv
         mkdir -p cropped_videos
         for i in *.mp4; do
           ffmpeg -i "$i" -vf "crop=180:50:0:0" "cropped_videos/${i}"
         done

This will create a new folder within present working directory that contains cropped label videos, use videos from this directory when `Explore_Obj2`-labeled videos are required.

#### 4.2.b. Overlay Cropped `Explore_Obj2` Labels onto `Explore_Obj1` Videos

    1. Combine cropped `Explore_Obj2` labels and `Explore_Obj1`-labeled videos into single directory, save                                   `auto_overlay_object_on_scene.m` MATLAB code in same directory.
    2. Add `Explore_Obj#`-labeled video directory to MATLAB Path (as done in 3.2.2 - 3.2.3).
    3. Enter following command into MATLAB command window: 
         auto_overlay_object_on_scene
    4. Select labeled video directory when promted and allow script to run, a completion message will be displayed upon successful           execution.	 

NOTE: The behavior labels were sized to fit videos with dimensions of 380x380. If your video dimensions differ, steps 4.1 and 4.2.a         will vary slightly:
    4.1: within `auto_processAllVideosObj#.m` codes: modify `FontSize` and x-and-y coordinates listed within `frame = insertText` to           modify label size and position respectively. 
    4.2.a: modify crop dimensions to reflect changes in 4.1.


Please cite Mar et al 2025 (Update when published)

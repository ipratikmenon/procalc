PROCALC HYDRAULICS - launcher pack
===================================

Files (keep them in ONE folder)
  run_hydraulics.bat     double-click: text menu (Windows).  run_hydraulics.sh / .command for Linux / macOS
  run_hydraulics.py      the menu / command-line front end used by the .bat
  hmbbuilder.py          PRO/II dump  ->  HMB workbook   (replaces td_parser / hmb_proii_reader / comp_constants)
  hydraulics.py          the hydraulics engine (piping classes, units, stream map, PR databank inside)
  pipeline_input_template.xlsx   blank input workbook (Pipeline_Input, Stream_Composition, Pseudo_Components,
                                 Stream_Map, Flow_Cases, Instrument_Data, Notes, UNITS)

Needs: Python 3.9+ and the openpyxl package (the .bat installs it if it is missing).

1. Build the HMB and run the hydraulics from the menu
  - Copy the PRO/II dump (e.g. BTGOLD_HCU_CASE2.xlsx) and your input workbook into this folder.
  - Double-click run_hydraulics.bat.
      1  Run hydraulics       - pick the input workbook, then an HMB; a PRO/II dump in the list is built into
                                an HMB first ("PRO/II dump -> build the HMB first")
      2  Build the HMB from a PRO/II dump, then run hydraulics
      3  Build the HMB from a PRO/II dump only (writes HMB_<dump name>.xlsx, ~80 s for ~460 streams)
      4  Create a blank input template      5  Open this folder
  - The run asks for the HMB, the circuit (or ALL), the flash mode (isothermal / isenthalpic) and the output
    colours (coloured = original dark banners, default; light = Technip Energies light theme + TEN logo).
  - Choose 'p' in a list to type a full path to a file that lives elsewhere.

2. Command line (no questions)
  run_hydraulics.bat build DUMP.xlsx [HMB_OUT.xlsx] [--sheet CASE2] [--constants Constants.xlsx]
  run_hydraulics.bat run   INPUT.xlsx HMB.xlsx [CIRCUIT] [--case "Case 1"] [--isenthalpic] [--light]
  run_hydraulics.bat all   DUMP.xlsx INPUT.xlsx [CIRCUIT]        (build the HMB, then run)

3. What you get
  One output workbook per circuit per flow case:  hydraulics_flash_noiso_<circuit>_<stream>[_<case>].xlsx
  Flow cases (Normal 100 %, Maximum 110 %, Minimum 60 % ...) are set on the Flow_Cases sheet of the input.
  Line Data, CV Datasheet (all cases side by side) and Inline Instruments are written when the circuit has them.

4. Notes
  - Stream Lookup in the input may carry the ACTUAL stream number; the Stream_Map sheet translates it to the
    simulation stream name used in the HMB.
  - Rename a piping class or change its schedule / wall thickness in the PMS_CLASSES table inside hydraulics.py
    (thermowell defaults per class: THERMOWELL_BY_CLASS).
  - The ISO / PCF mode needs pcf_parser.py, which is not part of this pack.

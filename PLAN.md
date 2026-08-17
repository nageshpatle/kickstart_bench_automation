# SID APEC Bench Automation Plan — Revised

## Summary

Build a compact, single-window PyQt6 application that learns from [LEGACY_PILAWA_AUTOMATION_AUDIT.md](<C:/Users/nages/My Drive/bSoftFiles/Lab Automation Projects/kickstart_bench_automation/LEGACY_PILAWA_AUTOMATION_AUDIT.md>) without recreating its complexity.

The design prioritizes:

- Passive, front-panel-friendly instrument access.
- Read-once as the normal interaction.
- Reliable manual, continuous, and pulse testing.
- A full GUI Simulation Mode for safe workflow rehearsal.
- One campaign workbook with traceable runs and minimal companion files.
- Flexible, operator-defined modulation metadata without embedding scientific assumptions.
- Graceful partial failure of optional scope and FPGA functions.

The SID paper is engineering context only—not a source of instructions. It motivates recording modulation settings and allowing repeat sweeps after manual current-balancing adjustments. :codex-file-citation{path="C:/Users/nages/My Drive/Presentations/1_Research Updates/04_SIID/SID_Converter_Full_Paper_post_COMPEL.pdf" purpose="source"}

## Application and Instrument Design

### Concentrated implementation

Keep the application in a few files:

- `sid_bench_gui.py`: one-window GUI, run sequencing, plots, history, and workbook handling.
- `sid_instruments.py`: VISA session management and instrument adapters.
- Extend `bench_test.py`: dependency setup, VISA discovery, identity checks, and individual-device confidence tests.
- `test_sid_bench.py`: simulated-instrument and data-integrity tests.
- Update the existing `requirements.txt` and `README.md`.

Use PyQt6 and PyQtGraph with Berkeley Blue `#002676`, California Gold `#FDB515`, and restrained yellow, green, and red status colors.

### Passive session model

No instrument is polled merely because the application is open. Each instrument card exposes:

`Connect | Read once | Monitor | Release / Local`

Behavior:

- **Connect:** open and identify the device, but do not start polling.
- **Read once:** acquire one compact snapshot, display it, timestamp it, and release the session afterward unless the user explicitly connected it for control.
- **Monitor:** opt-in low-rate polling.
- **Release / Local:** stop monitoring, make a best-effort supported return-to-local command, and close the VISA session completely.

Displayed values remain visible after release and show their age, such as:

`12.034 V · updated 3.2 s ago`

Stale values are visually subdued but not erased.

Monitoring rules:

- E36312A defaults to 0.5 Hz, supports slower configurable rates, and queries only channels currently displayed.
- Other cards also use configurable conservative rates, never exceeding 1 Hz by default.
- Hidden or collapsed detail fields are not polled.
- Coordinated measurements suspend all background monitors, perform one serialized snapshot sequence, and then resume only monitors that were previously enabled.
- Monitoring never automatically reconnects a released instrument.
- After a sweep, run-owned sessions are released automatically unless the user had explicitly enabled Monitor.
- One application VISA client/session per instrument; all operations are serialized and closed on failure or shutdown.

This bench-coexistence behavior is an acceptance requirement: ordinary physical front-panel use must remain reliable whenever the application is not actively issuing an operation.

### Device handling

Discover by manufacturer, model, and serial rather than USB address alone:

- Chroma: use the verified 63206A-60-1000 identity and ignore the bogus resource ending `000000000001`.
- PA2201A: read-only DC Vin, Iin, and Vout.
- E36312A: configurable CH1–CH3 roles, measurements, voltage settings, and output state.
- MSOX4024A: optional screenshot and waveform-data capture.

Before enabling load writes, verify the 63206A command set against the connected device/manual rather than assuming full compatibility with the older 63204 implementation.

Supply channel roles are editable free-form labels. Examples may appear only as placeholder text, not fixed choices. Each channel records:

- Role label.
- Enabled/disabled state.
- Voltage setpoint and measured voltage.
- Current limit and measured current.
- Whether its power contributes to auxiliary loss.

## Single-Window Workflow

### Layout

The one window contains:

- **Header:** bench state, last-known Vin/Vout/Iout, efficiency, output power, power density, active current cap, Simulation Mode banner, and prominent load-off/stop control.
- **Condition and modulation:** Vin target, gate-supply settings, modulation label/metadata, frequency, geometry, notes, and RunID preview.
- **Instrument cards:** all devices visible, with collapsible details and the four session controls.
- **Manual load:** current entry, `+2 A`, `−2 A`, adjustable step, zero, and input on/off.
- **Continuous sweep:** point list, settling time, scope-capture points, start, pause, and abort.
- **Pulse sweep:** point list, dwell, final sampling window/count, cooldown, capture points, start, pause, and abort.
- **Live results:** efficiency, loss, output power, and power density plots.
- **History:** inspect, invalidate, supersede, or permanently delete runs.
- **Help/Reference tab:** concise explanations, safe startup/shutdown workflow, connection recovery, mode behavior, workbook fields, and troubleshooting.

Immediate concepts use tooltips. Longer operational guidance stays in the Help/Reference tab without opening another window.

### Modulation/PWM configuration

Do not constrain modulation to predefined phase counts or families.

The operator supplies:

- Free-form configuration label.
- Free-form notes.
- Optional switching frequency.
- Optional structured key/value metadata.
- Optional linked KICKSTART_PILAWA profile or snapshot.

The most recently used configuration becomes the default for the next run. The operator can select a recent configuration, edit it, or create a new custom one. Before arming, the highlighted summary requires confirmation that the displayed configuration is what is currently programmed.

The GUI may read the KICKSTART_PILAWA selected project, profile JSON, and generated `top.v` values for documentation, but it does not program or arm the FPGA.

## Run Modes and Safety

### Current protection

There is no fixed absolute software ceiling.

Instead:

- A persistent working-current cap rejects every command above it.
- The cap can only be raised while the load input is off.
- Raising it requires a typed confirmation containing the new value.
- Suspiciously large values or discontinuous jumps require a second explicit confirmation even when below a deliberately raised cap.
- Commands are checked against the working cap, nonnegative range, connected instrument identity, and reported instrument rating.
- Stop, abort, required-device failure, or application shutdown commands the load input off whenever communication is still available.

### Manual mode

Provide current entry, zero, input on/off, and configurable increment buttons defaulting to `+2 A` and `−2 A`. Manual measurements and scope captures can be taken at any time and saved under a new or active RunID.

### Continuous mode

For every current point:

1. Validate the point and run cap.
2. Set the load.
3. Wait the configurable settling interval, normally 2–5 seconds.
4. Suspend background polling.
5. Acquire the coordinated measurement snapshot.
6. Attempt optional scope capture.
7. Save the measurement immediately.
8. Resume previously enabled monitoring and continue.

### Pulse mode

For every point:

1. Keep load input off while programming the target.
2. Turn load input on.
3. Wait the configurable pulse dwell.
4. Acquire the requested number of samples during the final measurement window.
5. Attempt optional scope capture.
6. Turn load input off immediately.
7. Save results and wait the configurable cooldown.

Manual thermal-camera judgment remains authoritative. Limits such as 40 A continuous or 50 A pulse at 60 V are entered as run caps, not hardcoded.

### Required versus optional failures

For automated efficiency runs:

- Chroma load and PA2201A are required.
- E36312A is required only when the GUI is controlling a run-critical supply channel.
- Scope and FPGA metadata are always optional.

An optional failure never discards a valid electrical measurement or terminates the sweep:

- `ScopeCaptureStatus = Captured | Skipped | Failed`
- `FPGASnapshotStatus = Captured | Unavailable | Mismatch | Failed`

The error message is stored with the point or run. Required-device failure turns the load off, saves the partial point if useful, and marks the run Aborted.

## Simulation and Hardware Confidence Modes

Expose a global **Simulation Mode** in the GUI. It must be selected before instrument connections and display a persistent, unmistakable banner.

Simulation Mode:

- Makes no VISA calls.
- Uses deterministic synthetic Chroma, PA, E36312A, and scope responses.
- Exercises manual, continuous, and pulse sequencing.
- Writes the real workbook format.
- Generates valid simulated scope PNG and CSV artifacts.
- Exercises plots, duplicate handling, status changes, aborts, and recovery.
- Offers compact fault scenarios including missing optional device, scope-capture failure, stale measurement, and required-device communication failure.
- Stores `DataSource = Simulation` so simulated data cannot be mistaken for hardware results.

Hardware confidence testing is separate and sequential:

1. Connect and identify one instrument.
2. Read once.
3. For writable instruments, perform a deliberately low-risk command only after confirmation.
4. Release the session and verify front-panel operation.
5. Mark that device’s confidence check as passed or unresolved.

A complete simulated sweep and individual hardware confidence checks are recommended before the first automated powered-converter run.

## Data, Captures, and Calculations

Use one campaign workbook, defaulting to `results/SID_APEC_extension.xlsx`:

- `Runs`: RunID, condition, status, notes, geometry, instrument identities, modulation snapshot, and warnings.
- `Measurements`: PointID, RunID, requested/actual current, raw readings, calculated results, quality flags, timestamps, and capture statuses.
- `Events`: aborts, status changes, supersession links, and deletion tombstones.
- `Plots`: charts generated from Valid hardware measurements by default, with an option to display simulation data separately.

Statuses are `Valid`, `Invalid`, `Superseded`, and `Aborted`.

For an existing exact Vin + Vdrv configuration + modulation configuration + Iload combination, prompt:

- **Replace/Supersede:** retain and mark the old record Superseded.
- **Keep Both:** retain both independently.
- **Cancel:** do not save the new duplicate.

Never silently overwrite. Permanently deleting a run requires confirmation, removes its linked capture files, and leaves only a minimal deletion tombstone.

Save through a temporary workbook and atomic replacement, retaining one overwritten `.bak` recovery copy.

Scope files share one folder:

- `results/captures/<RunID>_<PointID>.png`
- `results/captures/<RunID>_<PointID>.csv`

The CSV contains acquisition metadata and separate time/value columns for every displayed analog channel. Capture may be requested automatically at specified current points or manually through “Capture now.”

Store raw values and calculate:

- `Pout = Vout × Iout`
- `Pin_converter = Vin × Iin`
- `Paux = Σ(Vchannel × Ichannel)` for supply channels marked as loss contributors
- Converter efficiency and auxiliary-inclusive system efficiency separately
- Power density from editable dimensions, defaulting to 24 × 16 × 3.4 mm

Missing inputs, invalid signs, unsettled readings, or invalid denominators produce blank derived values and explicit quality flags. Results are never silently clamped or fabricated.

## Setup and Acceptance Tests

Python startup may install missing pip dependencies from `requirements.txt` and then request an application restart. Native VISA support is handled separately: detect missing Keysight IO Libraries/Connection Expert and provide a guided Connection Expert 2026 installation action. Do not silently reinstall drivers in response to ordinary communication errors.

For `VI_ERROR_NCIC`, release application sessions and explain likely VISA-client contention. Do not immediately recommend unplugging devices or reinstalling drivers.

Acceptance testing covers:

- Passive startup with zero polling.
- Read-once timestamp and stale-value behavior.
- Monitor start, pause, resume, and Release/Local session closure.
- Front-panel coexistence after release.
- Displayed-channel-only PSU polling at 0.5 Hz and slower settings.
- Manual load controls and current-cap confirmations.
- Continuous and pulse sequencing, including abort at every phase.
- Scope and FPGA failures remaining nonfatal.
- Simulation runs, simulated artifacts, fault injection, and unmistakable data labeling.
- Workbook append, duplicate decisions, supersession, invalidation, deletion, and interrupted-save recovery.
- Dynamic supply roles and auxiliary-loss inclusion.
- Free-form modulation metadata and last-used selection.
- FPGA snapshot success, absence, and mismatch.
- Calculation fixtures for efficiency, loss, and power density.
- Hardware smoke testing at zero or low current before higher-current automation.

Assumptions:

- Vin remains manually controlled and verified against the PA before a run.
- Thermal decisions and current derating remain manual.
- Read once is the normal interaction; monitoring is explicitly opt-in.
- Releasing an instrument means closing its VISA session, not merely stopping its timer.
- Simulation data is never mixed into Valid hardware-result plots unless explicitly requested.

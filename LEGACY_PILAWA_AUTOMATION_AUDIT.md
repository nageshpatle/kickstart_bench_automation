# Legacy Pilawa Lab Bench Automation Audit

## Executive crux

The legacy Pilawa implementation is a valuable record of real bench workflows, instrument commands, experimental presets, and result formats. It is much more capable than the current baseline: it coordinates sources, gate-drive rails, power meters, electronic loads, and oscilloscopes; performs continuous and pulsed efficiency sweeps; plots results live; saves configurations; and has produced a large body of experimental data.

It should **not** be used as the software foundation of the new implementation. It is a research prototype that grew by copying and modifying whole GUI scripts for individual experiments. Hardware access, GUI behavior, test sequencing, calculations, and file writing are mixed together; failures are frequently hidden; safety is mostly procedural; and several code paths are stale, inconsistent, or broken on current Python.

The right strategy is:

1. Preserve the legacy system as a read-only reference.
2. Extract and verify its useful SCPI commands, setup semantics, test sequences, and measurement formulas instrument by instrument.
3. Rebuild those capabilities behind a safe, testable driver and orchestration layer.
4. Add write/control features only after the read-only monitor and measurement provenance are trustworthy.

The current project has the opposite profile: it is a clean, intentionally read-only connectivity baseline, but almost all production functionality is still missing. Its three driver modules and two test recipes are placeholders, and its live monitor only measures the E36312A. The immediate goal should be a robust read-only monitor for all three present instruments, followed by guarded control and then automated tests.

## Audit scope and evidence

Legacy source inspected:

`../Remote control Pilawa group_v1.1/Remote control Pilawa group_v1.1`

Current source inspected:

`./`

The legacy folder contains:

- 48 Python source files, including 14 copied/forked Tkinter GUI variants.
- A reusable `pilawa_package` with serial, socket, source, load, meter, power-supply, function-generator, and source-meter classes.
- 21 text setup presets for named converters and experiments.
- Two copies of an eight-slide operating guide.
- 1,955 CSV result files, 22 PNG images, five XLSX files, and additional archived artifacts.
- GUI variants spanning the original bench, 25 A, 600 A, 1,000 A, 1,200 A, and 2,200 A load configurations, plus experiment-specific switched-bus variants.

The archived results prove that the application was used extensively, but they also expose data-quality debt: 1,928 of the 1,955 CSV files begin with an artificial all-zero data row created by the array initialization pattern. Nine distinct CSV header variants exist, including manually extended headers.

Representative evidence locations:

- The intended operator workflow is documented in `Instruction of remote control GUI_pilawa group_20201121.pptx`.
- The feature-rich later workflow is in `remote_control_pilawagroup_GUI_1000AELoad_noVin_swBusVPD_transient.py`.
- The reusable instrument catalog is exported by `pilawa_instruments.py:10-40`.
- The electronic-load dynamic mode is implemented in `pilawa_package/electronic_loads/chroma_63204.py:37-108`.
- Scope screenshots and measurements are handled in `reados.py:12-104`; the more complete waveform scaling/export implementation is in `InfiniiVision_Analog_Waveform_Grab_Elegant.py:84-547`.
- The current baseline and its safety intent are described in `README.md:1-45`.

## Complete legacy feature inventory

### Connectivity and transports

- Prologix GPIB-over-serial transport with address selection, manual reads, EOI/EOS configuration, and multi-device triggering.
- Direct TCP socket transport for LAN instruments.
- Read-only microcontroller serial interface.
- Direct PyVISA access for USB instruments such as the E3631x gate-drive supply and Keysight oscilloscope.
- Configurable serial port, baud rate, timeout, GPIB address, VISA address, and debug logging in portions of the driver layer.

### Instrument library

| Category | Legacy support | Exposed capabilities |
|---|---|---|
| Electronic loads | Agilent 6060B, Chroma 63204 | Mode/range, static setpoint, slew/transient settings, current/voltage readback, load on/off, dynamic pulse parameters |
| Function generator | Agilent 33250A | Function/frequency/amplitude/offset setup, output on/off, sweeps, bursts, arbitrary waveform programming, error query |
| DMM / DAQ | Agilent 34401A, 34410A, 34461A, 34970A; Fluke 45 | Triggering, single and buffered acquisition, continuous acquisition, scan lists, data count/clear, error query |
| Power analyzers | Yokogawa WT310, WT3000 | RMS/range/rate configuration, V/I/P sample retrieval, storage control, stored-record retrieval, error query |
| Power supplies | Agilent 6030A, 6632A, 6674A; GW Instek PSU-6025; MagnaPower XR; Xantrex XHR4025 | Voltage/current set, readback, output activation/deactivation, default state, limited error handling |
| Source meter | Keithley 2400 | Voltage/current source, compliance, linear sweep, list mode, buffered reads, activation/deactivation, errors |
| Oscilloscope | Keysight InfiniiVision family | Screenshot capture, configured measurements, active-channel detection, binary waveform transfer, preamble scaling, time-axis reconstruction, CSV export |
| Auxiliary rails | Keysight E3631x via `vdrive.py` | Per-channel voltage/current set, measurement, and per-channel output on/off |

This breadth is worth preserving as a reference catalog. It is not evidence that every driver is currently correct or compatible.

### Bench control

- Set input voltage and input current limit on a bench supply.
- Ramp source voltage in small steps to reduce inrush current (`powerSupply.py:6-15`).
- Set one or more auxiliary/gate-drive rails and their current limits.
- Address one or several electronic loads and distribute total current between selected load channels.
- Update setpoints separately from output enable state.
- Turn the configured source, auxiliary rails, and loads on or off from the GUI.
- Read live source, drive, and load measurements where the selected hardware path implements them.
- Support multiple high-current bench arrangements through specialized GUI variants.

### Measurement and efficiency calculation

- Acquire voltage, current, and power tuples from Yokogawa power analyzers.
- Extract multiple input/output elements from WT3000 comma-separated readings.
- Read electronic-load current independently of the power analyzer.
- Calculate and display a single-point efficiency measurement.
- Accumulate manual measurement points before writing a CSV.
- Plot output current versus efficiency in an embedded Matplotlib graph.
- Record nominal `Pin` and `Pout` along with voltage/current and auxiliary-rail measurements.

### Automated test workflows

- Increasing or decreasing output-current sweeps.
- User-configurable start, stop, step, delay, output filename, and selected load channels.
- Continuous mode: set the load, wait for settling, then measure.
- Pulse mode: set the load, wait, measure, return the load to zero, and wait between points.
- Chroma dynamic-current pulse configuration with high/low currents, rise/fall settings, and high/low durations.
- Manual single-point measurement and manual export.
- Experiment-specific presets for named converters, ratios, switching conditions, pulse tests, and high-current ranges.

### Results and reproducibility features

- Timestamped CSV filenames for automated sweeps.
- CSV columns for output current, input/output voltage/current/power, efficiency, and gate-drive rails.
- Live efficiency plot during a sweep.
- Save and recall text setup files.
- Named setup files that act as practical experiment recipes.
- Oscilloscope screenshot capture to PNG.
- Oscilloscope waveform export with a physical time axis and scaled channel values.
- A large archive of real results that can be used to validate a future importer and plotting layer.

### Operator interface

- One-page Tkinter control panel.
- Editable source, load, auxiliary-rail, sweep, delay, load-channel, and filename fields.
- Separate Update, On, Off, Measure, Manual Efficiency, Sweep, Pulse Trigger, Save Scope, Save Setup, and Recall Setup actions.
- Visible measured input/output voltage/current and efficiency.
- Continuous-versus-pulse mode selector.
- Embedded efficiency graph.

## What the legacy implementation did well

### It captured real experimental workflows

The most valuable artifact is not the GUI code; it is the accumulated lab knowledge. The setup presets show realistic converter-specific ranges, settle times, pulse/continuous choices, multi-load arrangements, and naming conventions. Those should inform recipe design and hardware-in-the-loop acceptance tests.

### It separated some instrument-specific knowledge

The `pilawa_package` structure is directionally correct. Common operations such as `setValue`, `readCurrent`, `activate`, `deactivate`, `getLastSample`, and `checkError` are kept in device-specific modules rather than duplicating every SCPI command in the GUI.

### It recognized important bench behaviors

- Input-voltage ramping reduces inrush risk.
- Setpoint update is separate from output enable.
- Continuous and pulsed tests have different settling behavior.
- Multiple load channels can share high output current.
- Auxiliary-rail consumption can be relevant to system efficiency.
- Scope screenshots and raw waveforms belong with numerical results.

### Its scope waveform handling is technically substantial

The Keysight example-based implementation detects active channels, obtains waveform preambles, chooses transfer formats, reconstructs time, scales raw ADC values, and writes unit-bearing CSV columns. This is a much better starting reference than a simplistic raw `:WAV:DATA?` dump.

### It generated useful, inspectable artifacts

CSV, PNG, text presets, and timestamped names are simple and durable. The result archive is human-readable and does not depend on a database or proprietary viewer.

## Missing, unsafe, or unreliable behavior

### P0 — hardware safety and deterministic shutdown

1. **No formal bench state machine.** There is no enforced progression such as Disconnected → Safe → Armed → Running → Fault. Any button can call hardware methods based on mutable GUI fields.
2. **No centralized emergency shutdown.** GUI close only closes the Prologix serial port (`...swBusVPD_transient.py:854-856`); it does not prove that the load, source, and gate-drive outputs are off.
3. **No exception-safe all-off path.** A timeout, parse error, file error, or user interrupt during a sweep can leave the latest load/source state active.
4. **No watchdog or heartbeat.** Loss of the GUI, USB, serial, or GPIB connection is not tied to a safe hardware response.
5. **Limits are incomplete and scattered.** Some branches use `min()` to cap a command, but there is no single validated envelope for voltage, current, power, slew, duration, temperature, or allowed instrument/channel combinations.
6. **No readback verification after writes.** Setpoints and output states are generally assumed rather than queried and confirmed.
7. **No preflight topology check.** The program does not verify that the expected model/serial number is at each address before enabling hardware.
8. **No remote/local policy.** The code does not manage front-panel lockout or explicitly return instruments to local control. Reopening a VISA session for each E3631x operation in `vdrive.py` creates unnecessary control churn, while the current monitor holds a persistent remote session and polls too aggressively.
9. **Turn-off logic is not trustworthy in all variants.** A representative variant instantiates an Agilent 6060B driver at the configured load address before also sending Chroma off commands (`...swBusVPD_transient.py:426-444`). That is stale copied logic and could send the wrong command dialect.
10. **The GUI is blocking.** Sweeps and sleeps run on the Tkinter event thread, so the interface can freeze precisely when an operator needs Stop or Off.

Required improvement: one authoritative `safe_shutdown(reason)` sequence must attempt load-off first, then source/auxiliary off or ramp-down as appropriate, verify states, log every outcome, remain callable after partial initialization, and run from normal stop, fault, window close, Ctrl+C, and unhandled-exception paths.

### P0 — measurement validity and provenance

1. **Efficiency formulas vary by copied GUI.** Channel indices and inclusion of drive power are hard-coded per experiment rather than defined in a named measurement topology.
2. **Bad data is silently converted into plausible data.** A broad `except` sets efficiency to zero, and the result is clipped to `[0, 1]` (`...swBusVPD_transient.py:492-498`). This hides division, parsing, sign, wiring, and instrument errors.
3. **Sign is discarded.** Many readings are wrapped in `abs()`, preventing detection of reversed polarity, regeneration, wiring mistakes, or a wrong analyzer sign convention.
4. **Displayed/recorded power is internally inconsistent in some variants.** The example calculates efficiency from `Vout * load_current`, but records the WT3000 `Pout`, which can remain zero in archived files.
5. **No sample quality policy.** There is no configurable averaging window, standard deviation, outlier rule, stability threshold, overload check, or stale-reading detection.
6. **Artificial zero rows contaminate nearly every result.** Arrays are initialized with a zero and later appended; 1,928/1,955 archived CSVs consequently start with a fake point.
7. **No measurement provenance.** Results omit instrument model/serial/firmware, VISA/GPIB address, driver version, calibration date, range, aperture/integration time, averaging method, formula/version, operator, DUT identifier, and full applied configuration.
8. **No uncertainty or validity status.** A row has numbers but no `valid`, `warning`, `fault`, or reason fields.
9. **Timestamps are inconsistent.** A module-level timestamp is reused by manual export, while other actions generate a fresh timestamp. Collision/overwrite behavior is not uniformly controlled.

Required improvement: preserve raw signed readings; derive named quantities from an explicit measurement map; never clip or fabricate a result; write invalid rows with a reason; and save enough provenance to reproduce the calculation.

### P1 — architecture and maintainability

1. **Fourteen whole-GUI forks encode configuration as code.** Fixes must be repeated and behavior drifts between 600 A, 1,000 A, 1,200 A, 2,200 A, and experiment-specific copies.
2. **Hardware is opened at import time.** Merely importing a GUI can open COM4, initialize a meter, reset/configure hardware, and sleep.
3. **GUI, orchestration, calculations, drivers, and persistence are coupled.** There is no headless recipe runner or testable service layer.
4. **Addresses are hard-coded to particular benches and users.** Commented alternative addresses and absolute setup paths make accidental bench mismatches likely.
5. **Resource ownership is inconsistent.** Some drivers share one Prologix session; direct VISA helpers create a new resource manager/session for every query; cleanup is manual.
6. **Driver interfaces are inconsistent.** Similar operations use `turnon`, `activate`, `setOutputOn`, `setValue`, `readData`, `getLastSample`, and different return formats.
7. **Errors are printed, swallowed, or allowed to crash.** There is no typed error taxonomy, retry policy, event log, or fault escalation.
8. **No capability discovery.** The code assumes a selected driver supports the requested mode/range/command.
9. **Commented-out code is used as configuration/version control.** Large inactive branches obscure what the running experiment actually does.

Required improvement: one configurable application with transport adapters, typed instrument drivers, a bench coordinator, independent recipes, a result writer, and a GUI that calls those services.

### P1 — compatibility and code correctness

1. The original `remote_control_pilawagroup_GUI.py` does not parse because line 363 contains a stray backtick.
2. The scope waveform script documents Python 2.7, PyVISA 1.8, Windows 7, and `visa32.dll`; it needs a deliberate Python 3/64-bit port.
3. `socket_wrapper.py` sends and concatenates Python strings where modern sockets require bytes.
4. `serial_wrapper.py` returns bytes but compares some results to text strings.
5. `yokogawa_wt310.py` lowercases the mode, then uses identity comparison against uppercase `'RMS'` (`:11` and `:28`), which is logically incorrect.
6. Several modules depend on NumPy, SciPy, Matplotlib, PySerial, Tkinter, and PyVISA, but there is no legacy requirements/lock file or supported runtime declaration.
7. There are checked-in `.pyc`, `__pycache__`, IDE files, swap files, output data, and duplicate autosaved documentation.

Required improvement: port only verified logic into typed Python 3 code; do not make the legacy tree itself the production package.

### P1 — operator experience

1. No clear distinction between monitoring, armed control, active test, and faulted states.
2. No always-available Stop/E-stop control backed by a non-blocking worker.
3. No per-instrument remote/local indicator, last-update age, or query-rate display.
4. No structured preflight checklist or explicit confirmation of dangerous ranges.
5. No progress, estimated duration, current step, retry count, or reason for pause/fault.
6. Connection and parse failures can leave stale values visible without an obvious stale/error state.
7. A filename field can embed subdirectories, but directory creation and path validation are inconsistent.

### P2 — engineering quality and long-term lab use

- No automated unit, transcript, integration, or hardware-in-the-loop tests.
- No simulator/fake instruments for development away from the bench.
- No CI, formatting, static typing, packaging, or release/version process.
- No schema version for setup or result files.
- No migration/import tool for old presets and CSVs.
- No resume/checkpoint behavior after an interrupted long sweep.
- No immutable run manifest or audit/event trail.
- No calibration-expiry, maintenance, or instrument-health tracking.
- No role/permission boundary between monitoring and control.
- No formal support matrix linking driver versions to instrument firmware.

## Assessment of the current implementation

### Current strengths

- `bench_test.py` is intentionally read-only and closes every instrument session.
- The Keysight VISA DLL/PATH bootstrap is documented and automated.
- The three present instruments and their exact USB resources are explicit.
- Timeouts are bounded for connectivity checks.
- The README already states the correct future safety principles: outputs off by default, hard limits, explicit arming, exception-safe shutdown, and logging.
- The directory structure anticipates separate drivers and recipes.

### Current gaps

- `instruments/pa2201a.py`, `chroma63206a.py`, and `e36312a.py` contain empty classes.
- `tests/efficiency_sweep.py` and `tests/pulse_test.py` raise `NotImplementedError`.
- `live_monitor_gui.py` uses a generic instrument wrapper instead of the driver package.
- PA2201A and Chroma panels only connect and show `*IDN?`; their displayed measurement fields are never populated (`live_monitor_gui.py:312-326`).
- The PSU executes six measurement queries every 0.5 seconds (`live_monitor_gui.py:275-328`). This is too aggressive for comfortable front-panel use on the observed bench.
- Monitoring starts automatically, has no toggle, has no per-device rate control, and does not restore local/front-panel control.
- Query exceptions are silently converted to `None`; the GUI does not show the error, retry state, stale age, or disconnect.
- A failed initial connection is never retried.
- Shutdown closes sessions but does not join the polling worker or explicitly coordinate remote/local state.
- Configuration and addresses remain hard-coded.
- `requirements.txt` includes only PyVISA; no test/development tooling or project packaging exists.

## Recommended target architecture

### 1. Transport/session layer

- One shared VISA resource manager owned by the application.
- One session object per instrument with an explicit lifecycle.
- Serialized access per session, bounded timeouts, cancellable queries, and a small explicit retry policy for read-only operations.
- Model/serial validation immediately after `*IDN?`.
- Command/response transcript logging with sensitive or bulky binary data summarized.
- Optional remote/local hooks per instrument; never assume a generic `SYST:LOC` command works everywhere.
- No session or hardware activity at module import time.

### 2. Typed driver layer

Implement the present bench first:

- `PA2201A`: identity, status/error queue, configured channel/elements, signed V/I/P readings, overload/range status, read-only snapshot.
- `Chroma63206A`: identity, mode/range/setpoint/input-state queries, signed V/I/P readings, protected static-current control, input on/off, dynamic/pulse mode only after static control is certified.
- `E36312A`: identity, per-channel setpoint/output-state/readback, protected set operations, output on/off, and documented local-control behavior.

Each driver should return typed measurement objects with value, unit, timestamp, validity/status, and raw response. Control methods must validate requested values against both device capability and bench-configured safety limits.

Legacy drivers should be migrated only when a real experiment needs them and the exact model is available for verification.

### 3. Bench coordinator and safety state machine

Use one authoritative state model:

`DISCONNECTED → SAFE → ARMED → RUNNING → SAFE`

Any failure moves to `FAULT`, runs `safe_shutdown`, records verification results, and requires explicit operator acknowledgement before re-arming.

The coordinator owns:

- Connection/preflight and identity checks.
- Applied safety profile.
- Setpoint ramping and order of operations.
- Load/source enable sequence.
- Readback verification.
- Stop, fault, and shutdown behavior.
- Cancellation tokens so Stop remains responsive during settling and acquisition.
- An append-only event log.

### 4. Measurement topology and recipes

Define a configuration-level measurement map rather than hard-coded array indices. It must state which analyzer element or instrument supplies:

- `vin`, `iin`, `pin`
- `vout`, `iout`, `pout`
- each auxiliary/gate-drive input
- whether reported efficiency is power-stage-only or total-system efficiency

Recipes should be headless and independently testable:

- Read-only snapshot/monitor.
- Static operating point.
- Continuous efficiency sweep.
- Pulsed efficiency sweep.
- Electronic-load dynamic transient.
- Scope capture attached to a selected test point.

Every recipe validates the full point list before arming, predicts duration, exposes progress, checks stability/overload/fault criteria, and always exits through the coordinator.

### 5. Result model

Use one immutable directory per run:

```text
runs/<timestamp>_<slug>/
  manifest.json
  requested_config.json
  applied_config.json
  samples.csv
  events.jsonl
  plots/
  waveforms/
```

The manifest should include schema/software version, operator, DUT/run notes, start/end time, completion status, safety profile, instrument identity/firmware/address, measurement topology, formula version, sample/stability policy, and artifact hashes.

`samples.csv` should include requested versus measured setpoints, signed raw measurements, derived powers/efficiencies, validity, warning/fault reason, and timestamps. Invalid measurements remain invalid; they are never clipped into a valid range.

### 6. GUI behavior

- Monitoring defaults OFF until the operator connects, or uses a conservative 1–2 s period.
- Independent monitor toggle and rate per instrument.
- Minimum query set per refresh; avoid redundant channel selection and commands.
- Visible connected/remote/local/stale/fault status and last successful update time.
- All hardware work off the Tkinter thread; GUI updates return through `after()`.
- Persistent, always-enabled Stop control that cancels the recipe and invokes safe shutdown.
- Control widgets disabled unless the bench is ARMED and the safety profile accepts all entries.
- No silent exception handling; concise operator error plus detailed event log.

## Prioritized implementation roadmap

### Milestone 0 — freeze and extract legacy knowledge

- Keep the legacy folder read-only.
- Create a command/feature matrix for only PA2201A, Chroma 63206A, and E36312A.
- Verify every command against the correct programming manual and then on hardware in read-only mode.
- Define the present bench wiring/measurement topology and two efficiency definitions: power-stage and total-system if auxiliary power is measurable.
- Choose conservative bench safety limits with the user; do not infer them from old high-current presets.

Acceptance: all three instruments are positively identified; read commands, units, sign conventions, channel/element mapping, and remote/local behavior are documented and manually verified.

### Milestone 1 — safe driver and session foundation

- Implement session management, typed errors, identities, snapshots, event logging, and fake transports.
- Implement read-only drivers for all three current instruments.
- Add transcript-based unit tests and no-hardware simulator tests.

Acceptance: repeated read-only snapshots run for at least 30 minutes without resource leaks, UI lockup, stale values presented as live, or unacceptable front-panel interference.

### Milestone 2 — trustworthy monitor

- Refactor the GUI to call the drivers.
- Add monitor toggles, conservative rates, retry/reconnect behavior, stale indicators, and clean worker shutdown.
- Confirm the exact local-control restoration behavior of each device before enabling it.

Acceptance: all three panels show verified live readings; front panels remain usable according to the agreed bench behavior; stopping monitoring releases or localizes every instrument as designed.

### Milestone 3 — guarded static control

- Add the state machine, safety profiles, preflight, explicit arming, verified Chroma static-current control, and verified E36312A setpoint/output control.
- Implement and test `safe_shutdown` before exposing On/Off controls in the GUI.

Acceptance: injected timeout, disconnect, malformed response, GUI close, Ctrl+C, and worker exception all produce a logged safe shutdown attempt and verified final state or a loud unresolved-fault status.

### Milestone 4 — continuous efficiency sweep

- Implement topology-aware PA measurements, stability checks, averaging, signed/valid result handling, run manifests, CSV/event output, live plots, and cancellation.
- Import representative old CSVs to regression-test plotting and detect the legacy zero-row pattern.

Acceptance: a manually checked three-point sweep agrees with the instrument displays and an independent calculation within an agreed tolerance; no fabricated or clipped data is written.

### Milestone 5 — pulse/transient and scope integration

- Add Chroma dynamic mode only after static mode is proven.
- Port the useful InfiniiVision waveform scaling logic to current 64-bit Python/PyVISA.
- Associate every capture with an exact sweep point and manifest entry.

Acceptance: pulse timing/current levels are verified on the scope, cancellation is responsive, and scope artifacts are correctly scaled, timestamped, and linked to the point that produced them.

### Milestone 6 — optional legacy instrument migration

- Migrate a legacy driver only when its instrument and an active experiment require it.
- Use the same driver contract, identity verification, safety envelope, fake transcript, and hardware acceptance tests.
- Do not reproduce the old GUI-per-experiment pattern; express bench differences as versioned configuration and recipes.

## Non-negotiable acceptance criteria for the new system

- Outputs default off; monitoring cannot arm or enable anything.
- Every write is validated, logged, read back where possible, and attributable to a user action or recipe step.
- Stop remains responsive during waits and instrument I/O.
- Every exit/fault path attempts deterministic safe shutdown.
- No exception is converted into a plausible measurement.
- No result is clipped into physical plausibility.
- Signed raw values are retained; derived values identify their formula.
- Every row has a timestamp and validity/status.
- Every run records instrument identity, configuration, topology, software/schema version, and completion status.
- No addresses, serial numbers, limits, or measurement indices are embedded in GUI code.
- Read-only monitoring is certified before static control; static control is certified before automated sweeps; continuous sweeps are certified before pulse/transient tests.
- Hardware-free tests cover drivers, sequencing, cancellation, faults, persistence, and GUI/service integration.

## Final recommendation

Treat the legacy implementation as an **experimental knowledge base and regression corpus**, not as a codebase to modernize in place. Its recipes, SCPI clues, measurement mappings, ramp behavior, scope workflow, and archived results are highly useful. Its duplicated GUIs, global hardware sessions, silent error handling, weak shutdown behavior, and unversioned data model should not be carried forward.

For the present project, the next build target is not an efficiency sweep. It is a **low-rate, read-only, three-instrument monitor backed by real drivers, explicit validity, observable errors, and verified remote/local behavior**. Once that is stable, implement the safety state machine and static guarded control before porting any legacy automated test.

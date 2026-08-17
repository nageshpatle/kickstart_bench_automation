# SID APEC Bench Automation

Compact, single-window automation for SID converter efficiency and power-density testing. The program starts passively: it opens no VISA sessions and performs no polling until an operator requests it.

## First start

```powershell
Copy-Item bench_config.example.json bench_config.json
python bench_test.py --install
python bench_test.py --simulation
python sid_bench_gui.py
```

Edit the copied `bench_config.json` for this bench's VISA resources and optional FPGA checkout. The local configuration is intentionally ignored by Git so operator settings, hardware serials, and machine-specific paths are not published.

Use **Simulation Mode** for a complete dummy sweep before connecting hardware. Simulation writes the real workbook schema and clearly labels every row `DataSource = Simulation`.

For passive real-device discovery:

```powershell
python bench_test.py --discover
```

Native VISA support requires Keysight IO Libraries Suite / Connection Expert 2026. Python can install pip dependencies, but it does not silently replace system-level VISA drivers.

## Operator workflow

1. Close Interactive IO or any other continuously polling VISA program.
2. Test one device at a time with **Connect → Read once → Release**.
3. Verify the Chroma 63206A static command dialect at zero or low current before enabling hardware writes.
4. Enter Vin, switching frequency, an optional run/profile label, supply roles, and the maximum allowed load current.
5. Review the highlighted summary and confirm that it matches the physical and FPGA setup.
6. Start the continuous or pulse run. Use **LOAD OFF** whenever behavior is unexpected; during a run it also stops the sequence.

Read once is the default. Monitor is opt-in; E36312A defaults to a 2-second interval (0.5 Hz), only displayed channels are queried, and all monitors pause during coordinated sweep measurements. **Release closes the VISA session entirely and returns local control.**

Scope capture deliberately uses **STOP → save screen and displayed-channel CSV → RUN**. It does not use Single acquisition, so an absent trigger cannot leave the capture waiting indefinitely. Scope failures remain nonfatal to electrical measurements.

The compact Bench cards embed resized product-identification thumbnails sourced from [TruePoint's PA2201A listing](https://truepointlab.com/product/keysight-pa2201a-integravision-power-analyzer-2-channels-1-phase/), [Keysight E36312A](https://www.keysight.com/us/en/product/E36312A/80w-triple-output-power-supply-6v-5a-2x-25v-1a.html), [Transcat's Chroma 63206A listing](https://www.transcat.com/rent-chroma-63206a-60-1000-dc-electronic-load), and [Batronix's MSOX4024A listing](https://www.batronix.com/shop/oscilloscopes/Keysight-MSOX4024A.html). They are visual identification aids only.

## Files

- `sid_bench_gui.py` — GUI, sequencing, workbook, plots, history, and FPGA metadata snapshot.
- `sid_instruments.py` — real and simulated instruments with explicit session ownership.
- `bench_test.py` — dependency bootstrap, passive discovery, and confidence check.
- `test_sid_bench.py` — hardware-free regression tests.
- `bench_config.example.json` — safe starting template for local configuration.
- `bench_config.json` — generated/local addresses, working cap, and last-used operator settings; intentionally not committed.

Results append to `results/SID_APEC_extension.xlsx`. Scope artifacts share `results/captures/`; the workbook keeps their exact paths and capture status. One `.bak` workbook is retained during atomic saves.

The entire `results/` tree is intentionally ignored by Git. This prevents large scope CSV files, screenshots, workbook backups, and measured data from inflating repository history. Back up valuable campaign data separately; Git will not store or recover it.

## Important hardware boundary

The PA2201A channel mapping and Chroma command dialect must be checked on the actual bench before serious testing. Start from `bench_config.example.json`, then keep the verified mapping in the ignored local `bench_config.json`; no default is a claim that every firmware/wiring configuration uses the same mapping. Scope and FPGA snapshot failures are nonfatal. PA or load failure aborts the run and attempts load-off.

If `VI_ERROR_NCIC` appears, release sessions and close competing VISA clients first. Do not immediately unplug every device or reinstall drivers.

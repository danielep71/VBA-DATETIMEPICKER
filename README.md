<img width="1766" height="296" alt="performance_manager_banner2" src="https://github.com/user-attachments/assets/8e1d69a6-f867-489a-b388-6760457bc838" />

# Class Performance Manager

> High-precision timing, checkpoint reporting, and Excel runtime control for VBA

---

## Part of a larger framework

This component is part of the **Excel VBA Runtime Framework**.

Within that framework, `cPerformanceManager` acts as the **execution and performance engine**.

It provides the foundation for:

- performance instrumentation
- runtime control
- repeatable benchmarking
- structured checkpoint reporting
- Excel environment optimization

Framework home:

[Excel VBA Runtime Framework](https://github.com/danielep71/excel-vba-runtime-framework)

---

<img width="1536" height="1024" alt="cPM_Home_reduced" src="https://github.com/user-attachments/assets/c4137fcb-2886-4d38-9cb8-e3349112c258" />

## Overview

`cPerformanceManager` is a **high-precision timing and execution-control component for Excel VBA on Windows**.

It wraps multiple timing backends behind a single, session-bound interface and adds a practical runtime-control layer for Excel automation.

The class supports:

- precise elapsed-time measurement
- human-readable elapsed-time diagnostics
- formatting of an already measured elapsed value without taking a second timing sample
- benchmark-overhead estimation
- pause / wait helpers
- structured checkpoints and reporting
- shared Excel “time-waster” suppression

Importantly, the suppression features are **not limited to timed benchmarks**. They can also be used as a **general-purpose Excel/VBA performance aid** in procedures that do not measure elapsed time, reducing avoidable overhead from screen refresh, events, alerts, calculation-mode churn, and cursor updates during heavy operations.

This makes `cPerformanceManager` more than a timer utility: it is a **runtime execution controller** for structured and performance-aware VBA solutions.

---

## Why this exists

VBA’s native timing options are often not ideal for instrumentation and benchmarking:

- `Timer` has limited resolution and rolls over at midnight
- `Now()` is wall-clock based and is not a preferred monotonic benchmark source
- Windows exposes better monotonic and high-resolution counters, such as `QueryPerformanceCounter`

`cPerformanceManager` provides a consistent abstraction over those timing sources and complements them with execution and environment controls that are highly useful in real Excel/VBA projects.

---

## Core capabilities

- multiple timing methods behind one interface
- session-bound timing model
- low-overhead numeric elapsed-time measurement
- human-readable elapsed-time formatting
- formatting of an already measured elapsed value without taking a second timing sample
- benchmark-overhead measurement helpers
- timing/source diagnostics
- pause / wait helper
- structured checkpoint capture within one timing session
- machine-readable checkpoint export
- human-readable checkpoint reporting
- shared Excel “time-waster” suppression for both benchmarking and general Excel/VBA performance improvement
- strict-mode validation for safer usage

---

## What are “time-wasters”?

“Time-wasters” are Excel application behaviors that can degrade performance during execution, especially in heavy procedures or repeated loops.

Typical examples include:

- screen updating
- event firing
- display alerts
- automatic calculation churn
- cursor state changes

`cPerformanceManager` provides centralized control over these elements so they can be suppressed during intensive procedures and restored cleanly afterward.

This is useful both for:

- cleaner benchmark runs
- faster ordinary workbook automation, even when no elapsed-time measurement is being taken

---

## Structured checkpoints and reporting

A single elapsed-time value is often not enough.

`cPerformanceManager` also supports **named checkpoints** inside a timing session so you can break a workflow into meaningful measured phases such as:

- loading data
- building arrays
- writing formulas
- recalculating
- exporting results

For each checkpoint the class stores:

- sequence number
- checkpoint name
- optional note
- delta seconds since the previous checkpoint
- cumulative elapsed seconds since `StartTimer`
- timing method metadata
- optional run label

The report can then be exported as:

- a **2D array** through `ReportAsArray()`
- a **readable multiline text block** through `ReportAsText()`

---

## Repository contents

```text
/README.md
/src/cPerformanceManager.cls
/src/M_cPM_TIMEWASTERS.bas
/demo/M_cPM_DEMO.bas
/demo/M_cPM_USAGE_EXAMPLES.bas
/demo/M_DEMO_BUILDER.bas
/test/M_cPM_TEST.bas
```

### Required files

These two files are required for normal use:

- `src/cPerformanceManager.cls`
- `src/M_cPM_TIMEWASTERS.bas`

### Optional companion files

These files are optional but useful:

- `demo/M_cPM_USAGE_EXAMPLES.bas`  
  Compact usage examples and recommended integration patterns

- `demo/M_cPM_DEMO.bas`  
  Interactive demo-sheet builder and runnable demo scenarios

- `demo/M_DEMO_BUILDER.bas`  
  Shared worksheet/demo layout helpers used by the demo environment

- `test/M_cPM_TEST.bas`  
  Regression test harness covering timing, fallback, TW lifecycle, cleanup, and checkpoint/reporting behavior

---

## Typical use cases

### Performance benchmarking

Measure execution time with high precision using a consistent API and a preferred high-resolution backend.

### Large dataset processing

Suppress unnecessary Excel overhead during intensive procedures to improve runtime performance.

### Controlled execution environments

Run procedures under predictable application-state conditions, then restore the Excel environment cleanly.

### Workflow instrumentation

Capture named checkpoints and export structured delta/cumulative timing for multi-step procedures.

### General workbook performance improvement

Use shared “time-waster” suppression even when elapsed time is not being measured.

---

## Timer methods

| ID | Method | Notes |
|---:|---|---|
| 1 | `Timer` | Seconds since midnight; rolls over every 24 hours |
| 2 | `GetTickCount / GetTickCount64` | Milliseconds since boot; 32-bit path wraps at ~49.7 days |
| 3 | `timeGetTime` | Millisecond counter with 32-bit rollover semantics |
| 4 | `timeGetSystemTime (MMTIME / TIME_MS)` | Millisecond source treated with 32-bit rollover semantics |
| 5 | `QueryPerformanceCounter (QPC)` | Default and recommended high-resolution benchmark source |
| 6 | `Now() * 86400` | Wall-clock seconds; mainly useful for diagnostics |

---

## Requirements

- Microsoft Excel with VBA enabled
- Windows host environment for API-backed timing methods (`2..5`)
- `VBA7` / `Win64` conditional compilation support as required by the host
- required source files:
  - `cPerformanceManager.cls`
  - `M_cPM_TIMEWASTERS.bas`

---

## Installation

1. Open the target workbook, add-in, or VBA project
2. Open the VBA Editor (`ALT + F11`)
3. Import:
   - `src/cPerformanceManager.cls`
   - `src/M_cPM_TIMEWASTERS.bas`
4. Save as macro-enabled (`.xlsm` or `.xlam`)
5. Compile the project (`Debug` → `Compile VBAProject`)
6. Run a smoke test

Optional:

- import `demo/M_cPM_USAGE_EXAMPLES.bas`
- import `demo/M_cPM_DEMO.bas`
- import `demo/M_DEMO_BUILDER.bas`
- import `test/M_cPM_TEST.bas`

---

## Quick start

### 1) Basic timing (default QPC backend)

```vb
Sub Example_BasicTiming()

    Dim cPM As cPerformanceManager
    Dim ElapsedS As Double

    Set cPM = New cPerformanceManager

    cPM.StartTimer

    Range("A1:A10000").Value = 1

    ElapsedS = cPM.ElapsedSeconds
    Debug.Print "Elapsed seconds: " & Format$(ElapsedS, "0.000000000")

    cPM.ResetEnvironment
    Set cPM = Nothing

End Sub
```

### 2) Format an already measured elapsed value

```vb
Sub Example_FormatExistingElapsed()

    Dim cPM As cPerformanceManager
    Dim ElapsedS As Double
    Dim ElapsedT As String

    Set cPM = New cPerformanceManager

    cPM.StartTimer 5, False
    Application.Calculate

    ElapsedS = cPM.ElapsedSeconds
    ElapsedT = cPM.ElapsedTime(, ElapsedS)

    Debug.Print "Elapsed seconds: " & Format$(ElapsedS, "0.000000000")
    Debug.Print "Elapsed time   : " & ElapsedT

    cPM.ResetEnvironment
    Set cPM = Nothing

End Sub
```

### 3) Structured checkpoints

```vb
Sub Example_Checkpoints()

    Dim cPM As cPerformanceManager
    Dim ReportArr As Variant

    Set cPM = New cPerformanceManager

    cPM.StartTimer 5, False
    cPM.SetRunLabel "ImportWorkflow"

    Range("A1:A10000").Value = 1
    cPM.Checkpoint "LoadValues"

    Range("B1:B10000").Formula = "=ROW()"
    cPM.Checkpoint "WriteFormulas"

    Application.Calculate
    cPM.Checkpoint "Recalculate", "Full workbook calculation pass"

    Debug.Print cPM.ReportAsText

    ReportArr = cPM.ReportAsArray

    cPM.ResetEnvironment
    Set cPM = Nothing

End Sub
```

### 4) Improve performance without measuring elapsed time

```vb
Sub Example_PerformanceOnly()

    Dim cPM As cPerformanceManager

    Set cPM = New cPerformanceManager

    cPM.TW_Turn_OFF

    Range("A1:A50000").Formula = "=ROW()"
    Application.Calculate

    cPM.TW_Turn_ON
    cPM.ResetEnvironment
    Set cPM = Nothing

End Sub
```

### 5) Benchmark with Excel TW suppression

```vb
Sub Example_BenchmarkWithSuppression()

    Dim cPM As cPerformanceManager
    Dim ElapsedS As Double

    Set cPM = New cPerformanceManager

    cPM.TW_Turn_OFF
    cPM.StartTimer 5

    Range("A1:A50000").Formula = "=ROW()"

    ElapsedS = cPM.ElapsedSeconds
    Debug.Print "Elapsed seconds: " & Format$(ElapsedS, "0.000000000")

    cPM.TW_Turn_ON
    cPM.ResetEnvironment
    Set cPM = Nothing

End Sub
```

---

## Core public API

### Timing

- `StartTimer(Optional iMethod As Integer = 5, Optional AlignToNextTick As Boolean = False)`
- `ElapsedSeconds(Optional iMethod As Integer = 0) As Double`
- `ElapsedTime(Optional iMethod As Integer = 0, Optional ElapsedSecondsIn As Variant) As String`

### Session / state inspection

- `T1 As Double`
- `T2 As Double`
- `ET As Double`
- `ActiveMethodID As Integer`
- `HasActiveSession As Boolean`
- `MethodName(ByVal Idx As Integer) As String`
- `StrictMode As Boolean` (Get/Let)
- `RunLabel As String`

### Diagnostics / benchmarking

- `OverheadMeasurement_Seconds(Optional iMethod As Integer = 5, Optional Iterations As Long = 1000) As Double`
- `OverheadMeasurement_Text(Optional iMethod As Integer = 5, Optional Iterations As Long = 1000) As String`
- `Get_SystemTickInterval As String`
- `QPC_Get_SystemTickInterval As String`
- `QPC_FrequencyPerSecond As String`
- `QPC_FrequencyPerSecond_Value As Double`

### Execution control / environment

- `Pause(ByVal dSeconds As Double, Optional ByVal iMethod As Integer = 1)`
- `ResetEnvironment()`

### Shared time-waster control

- `TW_Turn_OFF(Optional ByVal Except As TW_Enum = TW_Enum.None)`
- `TW_Turn_ON()`
- `TW_IsActive As Boolean`
- `TW_ActiveSessionCount As Long`

### Checkpoints / reporting

- `SetRunLabel(ByVal RunLabel As String)`
- `ClearCheckpoints()`
- `Checkpoint(ByVal CheckpointName As String, Optional ByVal NoteText As String = vbNullString)`
- `CheckpointCount As Long`
- `ReportAsArray() As Variant`
- `ReportAsText() As String`

---

## `TW_Enum` flags

Use with `TW_Turn_OFF Except:=...` as bitmask flags:

- `TW_Enum.None`
- `TW_Enum.ScreenUpdating`
- `TW_Enum.EnableEvents`
- `TW_Enum.DisplayAlerts`
- `TW_Enum.Calculation`
- `TW_Enum.Cursor`

Example:

```vb
cPM.TW_Turn_OFF TW_Enum.ScreenUpdating Or TW_Enum.EnableEvents
```

---

## Strict mode behavior

- default: `StrictMode = True`
- strict mode raises on:
  - invalid timer method
  - `ElapsedSeconds` before `StartTimer`
  - explicit elapsed method mismatch versus active session
  - `SetRunLabel` after checkpoint capture has already begun
  - `Checkpoint` before `StartTimer`
- non-strict mode attempts fallback or coercion where supported

---

## Design notes

- Prefer method `5` (`QueryPerformanceCounter`) for benchmark-grade timing
- Use `ElapsedSeconds` for numeric logic and machine-readable results
- Use `ElapsedTime` for user-facing display
- When you already have a numeric elapsed value, prefer `ElapsedTime(, ElapsedSecondsIn)` to avoid taking a second timing sample
- Use checkpoints when a workflow needs phase-level timing rather than only one final elapsed value
- Always call `ResetEnvironment` in normal flows and in error-cleanup paths
- “Time-waster” suppression is useful both for cleaner benchmarks and for improving performance in ordinary Excel/VBA procedures, even when no elapsed-time measurement is being taken

---

## Running examples and tests

- Import `demo/M_cPM_USAGE_EXAMPLES.bas` for compact usage examples
- Import `demo/M_cPM_DEMO.bas` and `examples/M_DEMO_BUILDER.bas` for the interactive demo workbook surface

<img width="1917" height="915" alt="cPM Demo Control Panel" src="https://github.com/user-attachments/assets/efb40917-51d5-4494-87a5-bcd8fb2c97da" />


- Import `test/M_cPM_TEST.bas` and run:

```vb
Run_cPerformanceManager_RegressionSuite
```

<img width="1918" height="919" alt="cPM Test Results" src="https://github.com/user-attachments/assets/876bfa0d-c678-45a7-a925-80ec76febefb" />

The regression suite covers:

- constructor/default state
- method mapping and fallback behavior
- elapsed-time behavior across all timing backends
- formatted elapsed-time behavior
- pause methods
- overhead helpers and diagnostics
- TW lifecycle and cleanup
- checkpoint/reporting behavior

---

## Position in the framework

Within the **Excel VBA Runtime Framework**, `cPerformanceManager` is the component responsible for **execution performance, checkpoint instrumentation, and runtime environment control**.

It is intended to work alongside complementary components for:

- UI management
- event-driven interaction
- broader Excel application architecture

Framework home:

[Excel VBA Runtime Framework](https://github.com/danielep71/excel-vba-runtime-framework)

---

## Wiki

For additional examples, notes, and repository-level guidance, see the project wiki:

[cPerformanceManager Wiki](https://github.com/danielep71/VBA-Performance_Manager/wiki)

---

## License

This project is licensed under the terms in `LICENSE`.

---

## Author

Daniele Penza

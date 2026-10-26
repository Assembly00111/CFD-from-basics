# CFD in Excel

Three small 2D field solvers that run **entirely on Excel formulas** — no macros, no add-ins, no VBA.
Open the workbook, press **F9**, and watch the flow field develop in a colour-mapped grid.

| Workbook | Problem | Governing equation | Numerical method |
|---|---|---|---|
| `1_heat_conduction_2d.xlsx` | Transient heat conduction in a plate | ∂T/∂t = α∇²T | Explicit FTCS, 21×21 grid |
| `2_potential_flow_2d.xlsx` | Ideal flow over a bump in a channel | ∇²ψ = 0 | Gauss-Seidel, 41×21 grid, editable obstacle mask |
| `3_lid_driven_cavity_2d.xlsx` | Lid-driven cavity (incompressible Navier–Stokes) | ∇²ψ = −ω, Dω/Dt = ν∇²ω | Stream function–vorticity, explicit Euler, 33×33 grid |

## How it works

Every grid node is one spreadsheet cell. Each cell's formula references **itself** and its four neighbours, which creates intentional circular references. With Excel's **iterative calculation** turned on, every iteration re-evaluates all cells once, so **one Excel iteration = one solver step (or one relaxation sweep)**. Pressing F9 runs `Maximum Iterations` steps.

```
cell(i,j)  <-  f( cell(i,j), cell(i±1,j), cell(i,j±1) )     # evaluated 1000x per F9
```

A `Reset` switch on each `Inputs` sheet restores the initial condition (set to 1, then back to 0).

## Quick start

1. Download a workbook from [`workbooks/`](workbooks).
2. **Enable iterative calculation** (required):
   `File → Options → Formulas → ☑ Enable iterative calculation`, Maximum Iterations = `1000`, Maximum Change = `0.000000001`.
   > Excel takes this setting from the first workbook opened in a session. If you get a circular-reference warning, close all workbooks and reopen the CFD file first.
3. Press **F9** repeatedly until the monitors on the `Inputs` sheet stop changing.
4. Look at the field sheets (`T`, `Psi`, `Omega`, `u_vel`, `v_vel`, `Speed`) — they are heat-mapped with conditional formatting. On the `Psi` sheets, the dark bands are **streamlines**.

Colour convention: blue-on-yellow cells are inputs you may edit; everything else is calculated.

## The problems

### 1. Heat conduction

Edge temperatures are fixed (Dirichlet); the interior evolves with

```
T_new = T + Fo · (T_E + T_W + T_N + T_S − 4T),     Fo = α·dt/h² ≤ 0.25
```

The `Inputs` sheet shows `Fo` and a stability flag. The `Validation` sheet compares the centre-line temperature with the exact Fourier-series solution for a plate with one hot edge and three cold edges.

### 2. Potential flow over a bump

Solves Laplace's equation for the stream function ψ (u = ∂ψ/∂y, v = −∂ψ/∂x) with ψ = U·y at the inlet, ψ = 0 on the floor and on the solid, ψ = U·H on the ceiling and zero-gradient at the outlet. The `Mask` sheet holds the geometry — **type `1` into any cell to make it solid** (default: a semicircular bump, radius 0.3 m). Rules: solids must touch the floor, and the inlet/outlet columns must stay fluid.

Sanity check: clear the whole mask and you should recover the uniform stream ψ = U·y.

### 3. Lid-driven cavity

The classic CFD benchmark, in stream function–vorticity form so that **no pressure equation is needed**:

```
∇²ψ = −ω                                               (one Gauss-Seidel sweep per step)
∂ω/∂t + u ∂ω/∂x + v ∂ω/∂y = ν ∇²ω                      (explicit Euler, central differences)
u = ∂ψ/∂y,   v = −∂ψ/∂x
ω_wall = −2ψ_inner/h² − 2U_wall/h                      (Thom's formula; U_wall = 0 except the lid)
```

Inputs: Reynolds number, lid speed, `dt`, an upwind-blending factor θ (0 = central, 1 = first-order upwind) and streamline spacing. The sheet checks the three stability limits (`h²/4ν`, `4ν/U²`, `h/U`) for you. The `Validation` sheet plots the vertical-centre-line `u` profile against the Ghia, Ghia & Shin (1982) benchmark at Re = 100 and reports an RMS error.

## Verification status

- **Formula logic.** The exact formula strings stored in the workbooks were evaluated with a small Python interpreter and compared with an independent NumPy implementation of the same scheme: the cavity stream function agreed to ~1e-16.
- **Heat conduction.** Steady-state centre-line temperature is within ~0.09 °C of the exact series solution (plate at 0–100 °C), away from the heated edge itself.
- **Cavity (Re = 100, 33×33, NumPy reference of the same scheme).** ψ_min ≈ −0.100 vs −0.1034 (Ghia, 129×129); minimum centre-line u ≈ −0.204 vs −0.211. Expect a few percent error from the coarse grid and first-order wall vorticity.
- **Not verified:** I could not run these inside Microsoft Excel itself. LibreOffice's circular-reference solver does not handle these coupled iterations reliably, so use **Excel** (desktop). Please open an issue if you see different behaviour on your version.

## Limitations

- Grid sizes are fixed by the sheet layout. To change them, regenerate the workbooks (below).
- Explicit time stepping is conditionally stable — keep `dt` within the limits shown on `Inputs`.
- The cavity is reliable for Re ≲ 400 on the default grid; higher Re needs a finer grid and/or `θ > 0`.
- Potential flow is inviscid: no boundary layer, separation or wake.
- Iterative calculation recalculates on *every* edit, so the fields will keep evolving whenever you change a cell. Set `Reset = 1` to freeze them.
- If values blow up to error values, set `Reset = 1`, reduce `dt`, then set `Reset = 0`.

## Regenerate / change the grid

All workbooks are produced by one script (needs `openpyxl ≥ 3.1`):

```bash
pip install openpyxl
python generate_workbooks.py --out workbooks                    # defaults
python generate_workbooks.py --out workbooks --n-cavity 41 --n-heat 31 --nx-potential 61 --ny-potential 31
```

When you refine the cavity grid, reduce `dt` so it stays under `h²/(4ν)` and `h/U`.

## Repository layout

```
.
├── README.md
├── generate_workbooks.py
└── workbooks/
    ├── 1_heat_conduction_2d.xlsx
    ├── 2_potential_flow_2d.xlsx
    └── 3_lid_driven_cavity_2d.xlsx
```

## References

- U. Ghia, K. N. Ghia, C. T. Shin, "High-Re solutions for incompressible flow using the Navier–Stokes equations and a multigrid method", *J. Comput. Phys.* 48 (1982) 387–411.
- Thom, A. (1933), "The flow past circular cylinders at low speeds", *Proc. R. Soc. A* 141 — wall-vorticity boundary condition.
- J. H. Ferziger & M. Perić, *Computational Methods for Fluid Dynamics*, Springer.

## License

MIT

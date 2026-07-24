import numpy as np
import matplotlib.pyplot as plt
from pathlib import Path

npz_path = Path(
    "/home/gkaiwi58/FAM/FerroX/Exec/extracted_pz/"
    "MFIS_t_5_nomi_2.5_var_9/Pz_Phi_FE_all_voltage.npz"
)

data = np.load(
    npz_path,
    allow_pickle=True
)

# ==========================
# Load data
# ==========================

Pz_stack = data["Pz_stack"]
Phi_stack = data["Phi_stack"]
charge_stack = data["charge_stack"]

x_nm = data["x_nm"]

# FE-only z coordinates
z_nm = data["z_nm"]

# Full-device z coordinates
z_nm_full = data["z_nm_full"]

V_applied = data["V_applied"]

print("NPZ keys:", data.files)

print("Pz_stack shape     :", Pz_stack.shape)
print("Phi_stack shape    :", Phi_stack.shape)
print("charge_stack shape :", charge_stack.shape)

print("x_nm shape         :", x_nm.shape)
print("z_nm FE shape      :", z_nm.shape)
print("z_nm_full shape    :", z_nm_full.shape)

# ==========================
# Checks
# ==========================

assert Pz_stack.ndim == 3
assert Phi_stack.ndim == 3
assert charge_stack.ndim == 3

# Pz uses FE-only z coordinates
assert Pz_stack.shape[1] == len(x_nm)
assert Pz_stack.shape[2] == len(z_nm)

# Phi / charge use full-device z coordinates
assert Phi_stack.shape[1] == len(x_nm)
assert Phi_stack.shape[2] == len(z_nm_full)

assert charge_stack.shape[1] == len(x_nm)
assert charge_stack.shape[2] == len(z_nm_full)

assert Pz_stack.shape[0] == len(V_applied)
assert Phi_stack.shape[0] == len(V_applied)
assert charge_stack.shape[0] == len(V_applied)

# ==========================
# Select bias point
# ==========================

idx = 10

print("Selected stack index:", idx)
print("V_applied:", V_applied[idx])

P_slice = Pz_stack[idx]
Phi_slice = Phi_stack[idx]
charge_slice = charge_stack[idx]

print("Pz range:")
print(
    np.nanmin(P_slice),
    np.nanmax(P_slice)
)

print("Phi range:")
print(
    np.nanmin(Phi_slice),
    np.nanmax(Phi_slice)
)

print("Charge range:")
print(
    np.nanmin(charge_slice),
    np.nanmax(charge_slice)
)

# ==========================
# Plot Pz
# FE region only
# ==========================

plt.figure(figsize=(7, 3))

plt.pcolormesh(
    x_nm,
    z_nm,
    P_slice.T,
    shading="auto",
    cmap="RdBu_r",
)

plt.colorbar(
    label=r"Pz (C/m$^2$)"
)

plt.xlabel("x (nm)")
plt.ylabel("z (nm)")

plt.title(
    f"Pz_stack[{idx}], "
    f"V={V_applied[idx]:+.2f} V"
)

plt.savefig(
    f"check_Pz_{idx}.png",
    dpi=200,
    bbox_inches="tight",
)

plt.show()


# ==========================
# Plot Phi
# Full device
# ==========================

plt.figure(figsize=(7, 3))

plt.pcolormesh(
    x_nm,
    z_nm_full,
    Phi_slice.T,
    shading="auto",
    cmap="viridis",
)

plt.colorbar(
    label="Phi (V)"
)

plt.xlabel("x (nm)")
plt.ylabel("z (nm)")

plt.title(
    f"Phi_stack[{idx}], "
    f"V={V_applied[idx]:+.2f} V"
)

plt.savefig(
    f"check_Phi_{idx}.png",
    dpi=200,
    bbox_inches="tight",
)

plt.show()


# ==========================
# Plot charge
# Full device
# ==========================

plt.figure(figsize=(7, 3))

plt.pcolormesh(
    x_nm,
    z_nm_full,
    charge_slice.T,
    shading="auto",
    cmap="RdBu_r",
)

plt.colorbar(
    label="Charge"
)

plt.xlabel("x (nm)")
plt.ylabel("z (nm)")

plt.title(
    f"charge_stack[{idx}], "
    f"V={V_applied[idx]:+.2f} V"
)

plt.savefig(
    f"check_charge_{idx}.png",
    dpi=200,
    bbox_inches="tight",
)

plt.show()
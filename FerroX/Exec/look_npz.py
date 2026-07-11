import numpy as np
import matplotlib.pyplot as plt
from pathlib import Path

npz_path = Path("/home/gkaiwi58/FAM/FerroX/Exec/extracted_pz/MFIS_t_5_nomi/Pz_Phi_FE_all_voltage.npz")

data = np.load(npz_path, allow_pickle=True)

Pz_stack = data["Pz_stack"]
Phi_stack = data["Phi_stack"]

x_nm = data["x_nm"]
z_nm = data["z_nm"]

V_applied = data["V_applied"]
phi_available = data["phi_available"]

print("NPZ keys:", data.files)
print("Pz_stack shape:", Pz_stack.shape)
print("Phi_stack shape:", Phi_stack.shape)
print("x_nm shape:", x_nm.shape)
print("z_nm shape:", z_nm.shape)

assert Pz_stack.ndim == 3
assert Phi_stack.ndim == 3
assert Pz_stack.shape == Phi_stack.shape

assert Pz_stack.shape[1] == len(x_nm)
assert Pz_stack.shape[2] == len(z_nm)

assert Pz_stack.shape[0] == len(V_applied)
assert Phi_stack.shape[0] == len(phi_available)

idx = 10

print("Selected stack index:", idx)
print("V_applied:", V_applied[idx])
print("Phi available:", phi_available[idx])

P_slice = Pz_stack[idx]
Phi_slice = Phi_stack[idx]

print("P slice range:")
print(np.nanmin(P_slice), np.nanmax(P_slice))

print("Phi slice range:")
print(np.nanmin(Phi_slice), np.nanmax(Phi_slice))

# Plot Pz
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

# Plot Phi
plt.figure(figsize=(7, 3))

plt.pcolormesh(
    x_nm,
    z_nm,
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
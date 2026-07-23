import numpy as np

npz_path = "/home/gkaiwi58/FAM/FerroX/Exec/extracted_pz/MFIS_t_8_nomi_2.5_var_2/Pz_Phi_FE_all_voltage.npz"

with np.load(
    npz_path,
    allow_pickle=True,
) as data:

    for key in data.files:
        arr = data[key]

        print(
            f"{key:20s}",
            f"dtype={arr.dtype}",
            f"shape={arr.shape}",
        )
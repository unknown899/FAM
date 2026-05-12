In file included from ../../amrex/Src/Boundary/AMReX_LO_BCTYPES.H:4,
                 from ../../amrex/Src/Boundary/AMReX_lo_bctypes_mod.F90:2:
tmp_build_dir/s/3d.gnu.TPROF.MPI.OMP.EXE/AMReX_Config.H:48:2: error: #error libamrex was built with OpenMP
   48 | #error libamrex was built with OpenMP
      |  ^~~~~

















module amrex_lo_bctypes_module

  implicit none
  integer, parameter :: amrex_lo_dirichlet         = 101
  integer, parameter :: amrex_lo_neumann           = 102
  integer, parameter :: amrex_lo_reflect_odd       = 103
  integer, parameter :: amrex_lo_marshak           = 104
  integer, parameter :: amrex_lo_sanchez_pomraning = 105
  integer, parameter :: amrex_lo_inflow            = 106
  integer, parameter :: amrex_lo_inhomog_neumann   = 107
  integer, parameter :: amrex_lo_robin             = 108
  integer, parameter :: amrex_lo_symmetry          = 109
  integer, parameter :: amrex_lo_periodic          = 200
  integer, parameter :: amrex_lo_bogus             = 1729

end module amrex_lo_bctypes_module

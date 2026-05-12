In file included from ../../amrex/Src/Base/AMReX_omp_mod.F90:2:
tmp_build_dir/s/3d.gnu.TPROF.MPI.OMP.EXE/AMReX_Config.H:48:2: error: #error libamrex was built with OpenMP
   48 | #error libamrex was built with OpenMP
      |  ^~~~~













module amrex_omp_module

  implicit none

  integer, parameter :: amrex_omp_support = (_OPENMP)

  integer, external :: omp_get_num_threads
  integer, external :: omp_get_max_threads
  integer, external :: omp_get_thread_num
  logical, external :: omp_in_parallel

end module amrex_omp_module


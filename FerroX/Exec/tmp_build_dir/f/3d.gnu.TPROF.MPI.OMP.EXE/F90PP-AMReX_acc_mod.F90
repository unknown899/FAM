In file included from ../../amrex/Src/Base/AMReX_acc_mod.F90:2:
tmp_build_dir/s/3d.gnu.TPROF.MPI.OMP.EXE/AMReX_Config.H:48:2: error: #error libamrex was built with OpenMP
   48 | #error libamrex was built with OpenMP
      |  ^~~~~












module amrex_acc_module

  implicit none

  integer :: acc_stream

contains

  subroutine amrex_initialize_acc (id) bind(c,name='amrex_initialize_acc')
    integer, intent(in), value :: id
  end subroutine amrex_initialize_acc

  subroutine amrex_finalize_acc () bind(c,name='amrex_finalize_acc')
  end subroutine amrex_finalize_acc

  subroutine amrex_set_acc_stream (acc_stream_in) bind(c,name='amrex_set_acc_stream')

    implicit none

    integer, intent(in), value :: acc_stream_in

    ! Set the OpenACC stream (to be used with the async clause) to be consistente
    ! with the CUDA stream that AMReX is using.

    acc_stream = acc_stream_in

  end subroutine amrex_set_acc_stream

end module amrex_acc_module

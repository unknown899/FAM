/*
 * This file is part of FerroX.
 *
 * Contributor: Prabhat Kumar
 *
 */
#include <FerroXUtil.H>

using namespace amrex;


void FerroX_Util::Contains_sc(MultiFab& MaterialMask, bool& contains_SC)
{       /*
	amrex::Print() << "ENTER Contains_sc\n";

	amrex::Print() << "ok = "
		   << MaterialMask.ok() << "\n";

	amrex::Print() << "nComp = "
		   << MaterialMask.nComp() << "\n";

	amrex::Print() << "nBox = "
		   << MaterialMask.boxArray().size() << "\n";
	*/
	contains_SC = (MaterialMask.max(0) >= 2.0);
}

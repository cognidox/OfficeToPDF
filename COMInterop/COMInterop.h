#pragma once

using namespace System;

namespace COMInterop {
	public ref class OfficeApp
	{
		// TODO: Add your methods for this class here.
	public:
		DWORD GetProcessId(System::IntPtr application);
		DWORD GetProcessId2(System::IntPtr application);
	};
}

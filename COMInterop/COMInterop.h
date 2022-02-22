#pragma once

using namespace System;

namespace COMInterop {
	public ref class OfficeApp
	{
		// TODO: Add your methods for this class here.
	public:
		DWORD GetProcessId(System::IntPtr application);
	};
}

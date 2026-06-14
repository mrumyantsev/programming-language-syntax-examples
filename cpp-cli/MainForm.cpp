#include "MainForm.h"

using namespace System;
using namespace System::Threading;
using namespace System::Globalization;

using namespace MyProject;

[STAThreadAttribute]
int main(array<System::String ^> ^args) {
	// Setting ANSI Standards for all future/background threads.
	CultureInfo::DefaultThreadCurrentCulture = CultureInfo::InvariantCulture;
	CultureInfo::DefaultThreadCurrentUICulture = CultureInfo::InvariantCulture;

	// Setting ANSI Standards for the primary startup thread.
	Thread::CurrentThread->CurrentCulture = CultureInfo::InvariantCulture;
	Thread::CurrentThread->CurrentUICulture = CultureInfo::InvariantCulture;

	// Enabling visual styles.
	Application::EnableVisualStyles();
	Application::SetCompatibleTextRenderingDefault(false);

	// Creating of the main window and running it.
	Application::Run(gcnew MainForm());

	return 0;
}

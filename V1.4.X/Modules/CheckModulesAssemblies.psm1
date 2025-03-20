#region [Region 05 | WPF ASSEMBLIES]
# Loads required WPF assemblies for the GUI interface
Write-DebugMessage "Loading WPF assemblies."
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms
Write-DebugMessage "WPF assemblies loaded successfully."
#endregion
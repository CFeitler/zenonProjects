using Mono.Addins;

// Declares that this assembly is an add-in
[assembly: Addin("ButtonClickToCel", "1.0")]

// Declares that this add-in depends on the scada v1.0 add-in root
[assembly: AddinDependency("::scada", "1.0")]

[assembly: AddinName("ButtonClickToCel")]
[assembly: AddinDescription("This Service Engine AddIn Serivce reacts on button clicks and writes a cel entry.")]
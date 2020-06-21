# VbLoadDriver

.NET 4.0 app that loads a driver from the registry (optionally creating/updating registry values)

Heavily inspired by original C++ code here: https://github.com/TarlogicSecurity/EoPLoadDriver/

**USAGE:**

```VbLoadDriver.exe RegistryPath [DriverPath]```

**ARGUMENTS:**

```RegistryPath```

Path to registry key that contains parameters for this driver.
                Must start with HKLM (for HKEY_LOCAL_MACHINE) or HKU (for
                HKEY_USERS). If loading a new driver, specify a non-existent
                registry key here and the program will create the registry key
                and populate it with values that use the driver file specified
                by the DriverPath argument.

```DriverPath```

Full path to the driver file to be loaded, which will be
                written to the registry key specified by the RegistryPath
                argument. This argument is optional. Using it when specifying
                an existing registry key for RegistryPath will cause new values
                to be written to the registry key, overwriting existing values.

**EXAMPLES:**

VbLoadDriver.exe HKLM\System\CurrentControlSet\Services\SomeDriver

VbLoadDriver.exe HKU\S-1-5-21-1950...58-1040\System\CurrentControlSet\Services\MyNewDriver C:\NewDriver.sys

**NOTES:**

Your user account must already have the right to load drivers (SeLoadDriverPrivilege). This program just enables that existing privilege and makes use of it.

My initial tests seem to show that the system will only accept registry paths that are a subkey in the following location ```HKLM\System\CurrentControlSet\Services``` . Some versions of Windows will also accept that same location but in HKU (so ```HKEY_USERS\USER-SID-HERE\System\CurrentControlSet\Services```) so that's a lot more useful from a priv esc point of view, but I'm unsure of what patch or configuration changes this. HKU paths don't work on my Windows 7 and Windows 10 test machines, but do work on one Server 2016 machine I tested it on.

There's also some weird behaviour if the driver has already been loaded. You'll get errors like "cannot find the file specified" if you attempt to load the same driver from a different reg key and different file name/path. I can't seem to find a way to actually unload the driver other than a reboot either.

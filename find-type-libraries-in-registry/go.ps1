set-strictMode -version latest

new-psDrive -name HKCR -psProvider registry -root HKEY_CLASSES_ROOT

[Microsoft.Win32.RegistryKey] $regTypeLibs = get-item hkcr:/TypeLib

foreach ($guid in $regTypeLibs.GetSubKeyNames()) {

   $guid
   $regTypeLib = $regTypeLibs.OpenSubKey($guid)

   foreach ($version in $regTypeLib.GetSubKeyNames()) {

      $regVersion = $regTypeLib.OpenSubKey($version)
      $typeLibName = $regVersion.GetValue('')


      "   $version`: $typeLibName"

      foreach ($versionKey in $regVersion.GetSubKeyNames()) {
         "      $versionKey"

         $regVersionKey = $regVersion.OpenSubKey($versionKey)
         if ( -not @('0', '9', 'FLAGS', 'Flags', 'HELPDIR', 'HelpDir', '409').Contains($versionKey) ) {
           '!!!!!'
            return
         }

         if ($versionKey -match '^\d+$') { # typically a zero ('0', but other numbers are also seen)
#           $regNumber = $regVersionKey.OpenSubKey($versionKey)

            foreach ($winX in $regVersionKey.GetSubKeyNames()) {
               if (-not @('win32', 'Win32', 'win64', 'Win64').contains($winX)) {
                  "!!!!"
                  return

               }
               $regWinX = $regVersionKey.OpenSubKey($winX)
               $typeLibPath = $regWinX.GetValue('')
               "        $winX $typeLibPath" 
            }
         }
      }

   }


}

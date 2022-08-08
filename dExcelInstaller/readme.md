# Readme

## Testing the ∂Excel Installer

At some point it would be great to automate these tests, until such time they will have to be run manually.

1. From a clean base does it install a new instance of the ∂Excel addin. 
   1. Ensure ∂Excel is not in the *Excel Add-Ins* section of the *Developer* tab.
   2. Does it install when Excel is closed?
   3. Does it install when Excel is open?
   4. Does it purge the *Versions/Current* folder if the installer throws an exception/is interrupted?
2. From a base case where ∂Excel has been installed can it uninstall it.
   1. Does it uninstall when Excel is closed?
   2. Does it uninstall when Excel is open?
   3. Does it purge the contents of the *Versions/Current* folder if the installer throws an exception/is interrupted?
3. Does it update the local add-ins repository with those from the selected remote repo?

get-childitem -recurse | Where {$_.Name -eq "obj" -or $_.Name -eq "bin"} | Remove-Item -recurse

Set objShell = CreateObject("Wscript.Shell")
objShell.Run("devmgmt.msc")
wait 1
objShell.SendKeys"{TAB}"
objShell.SendKeys"{Down 2}"
objShell.SendKeys"{RIGHT}"
objShell.SendKeys"{Down}"
objShell.SendKeys"{ENTER}"
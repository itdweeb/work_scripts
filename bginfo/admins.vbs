SET colGroups = getObject("WinNT://.")
colGroups.filter = array("group")
FOR EACH objGroup IN colGroups
    IF objGroup.name = "Administrators" THEN
		FOR EACH objUser IN objGroup.members
			ECHO vbTab & objUser.name
		NEXT
	END IF
NEXT

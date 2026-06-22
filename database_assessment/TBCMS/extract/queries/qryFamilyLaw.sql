SELECT [Family Law - Divorce].*, tblCase.*
FROM tblCase RIGHT JOIN [Family Law - Divorce] ON tblCase.CaseID = [Family Law - Divorce].CaseID;
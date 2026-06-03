SELECT tblUsers.*, tblAccessType.*
FROM tblUsers INNER JOIN tblAccessType ON tblUsers.Access = tblAccessType.AccessType;
SELECT *
FROM Disposition
WHERE (((Disposition.Disposition) Like "n/p" And (Disposition.Disposition) Not Like "$" And (Disposition.Disposition) Not Like "/"));
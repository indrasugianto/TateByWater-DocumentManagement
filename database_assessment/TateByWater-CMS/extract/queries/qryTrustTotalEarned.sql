SELECT Sum([Trust Account].Credit) AS SumOfCredit, [Trust Account].TMatter, [Trust Account].CaseID
FROM [Trust Account]
GROUP BY [Trust Account].TMatter, [Trust Account].CaseID
HAVING ((([Trust Account].TMatter) Like "*Earned*"));
insertData = "15:45:51.444 [http-nio-8090-exec-2] DEBUG java.sql.PreparedStatement - {pstm-102832} Parameters: [323298968]"

subindex1 = insertData.rfind("[")
print(subindex1)
insertData = insertData[subindex1:].replace("[", "").replace("]", "")
print(insertData)
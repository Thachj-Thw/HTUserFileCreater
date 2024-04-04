from UserFileCreate import UserFileCreater


tool = UserFileCreater("test.xlsx")
tool.create()
print("created successfully", len(tool.list_license), "card")
tool.save()
import uno

localContext = uno.getComponentContext()
resolver = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext)
context = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
model = desktop.getCurrentComponent()
text = model.Text
cursor = text.createTextCursor()
text.insertString(cursor, 'Hello world!', 0)

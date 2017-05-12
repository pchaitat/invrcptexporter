REM  *****  BASIC  *****
REM
REM references:
REM   - https://ask.libreoffice.org/en/question/27312/create-a-pdf-for-every
REM     -sheet-named-after-the-sheet/

Sub export_invoices
	dim document as object
	dim dispatcher as object
	
	document = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

	path ="file:///home/ubuntu/tmp/sheet1b6yu22.pdf"
	Open path For Append As #1
	Close #1
	
	delranges
	
	invwtSheet = ThisComponent.Sheets.GetByIndex(1)
	ThisComponent.CurrentController.setActiveSheet(invwtSheet)
	
	dim args2(0) as new com.sun.star.beans.PropertyValue
	args2(0).Name = "PrintArea"
    args2(0).Value = "a1:h30"
	dispatcher.executeDispatch(document, ".uno:ChangePrintArea", "", 0, args2())
	
	dim args1(2) as new com.sun.star.beans.PropertyValue
	args1(0).Name = "URL"
	args1(0).Value = "file:///home/ubuntu/tmp/sheet1b6yu22.pdf"

	args1(1).Name = "FilterName"
	args1(1).Value = "calc_pdf_Export"

	dispatcher.executeDispatch(document, ".uno:ExportDirectToPDF", "", 0, args1())
End Sub

sub delranges
	rem from http://www.oooforum.org/forum/viewtopic.phtml?t=87886
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	rem loop to all my sheets and delete all printAreas
 	For index = 0 to Ubound(thisComponent.getSheets.getElementNames)
   		oSheet = ThisComponent.Sheets.getByIndex(index)
   		ThisComponent.CurrentController.setActiveSheet(oSheet)
   		document = ThisComponent.CurrentController.Frame
   		dispatcher.executeDispatch(document, ".uno:DeletePrintArea", "", 0,Array())   
	Next index
end sub

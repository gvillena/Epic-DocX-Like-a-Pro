Imports System
Imports System.IO
Imports TemplateEngine.Docx

Module Module1

    Sub Main()


        Console.WriteLine("")
        Console.WriteLine("INFORMES LIKE A PRO!")
        Console.WriteLine("")

        File.Delete("OutputTest.docx")
        File.Copy("InputTest.docx", "OutputTest.docx")

        Dim input As String = ""
        Dim numero As String = "0038"
        Dim filial As String = "LIMA"
        Dim fechaStr As String = Date.Now.ToString("dd/MM/yyyy")
        Dim bachiller As String = "JOHN SMITH"
        Dim tituloTesis As String = "LOREM IPSUM DOLOR SIT AMET, CONSECTETUER ADIPISCING ELIT. MAECENAS PORTTITOR CONGUE MASSA. FUSCE POSUERE, MAGNA SED PULVINAR ULTRICIES, PURUS LECTUS MALESUADA LIBERO, SIT AMET COMMODO MAGNA EROS QUIS URNA."
        Dim añoTesis As String = "2017"

        ' NUMERO INFORME
        Console.Write("NUMERO INFORME (0038): ")
        input = Console.ReadLine()
        If Not String.IsNullOrWhiteSpace(input) Then numero = input.PadLeft(input, "0")

        ' FILIAL
        Console.Write("FILIAL (LIMA): ")
        input = Console.ReadLine()
        filial = input.ToUpper()

        ' FECHA
        Console.Write("FECHA (" + Date.Now.ToString("dd/MM/yyyy") + "): ")
        input = Console.ReadLine()
        fechaStr = input

        ' BACHILLER 
        Console.Write("BACHILLER (APE. Y NOM.): ")
        input = Console.ReadLine()
        bachiller = input.ToUpper()

        ' TITULO TESIS
        Console.Write("TITULO DE TESIS (TEST): ")
        input = Console.ReadLine()
        tituloTesis = input.ToUpper()

        ' AÑO TESIS
        Console.Write("AÑO DE TESIS (TEST): ")
        input = Console.ReadLine()
        añoTesis = input

        Dim fcNumero = New FieldContent("Numero Informe", numero)
        Dim fcFilial = New FieldContent("Filial", filial)
        Dim fcFecha = New FieldContent("Fecha", fechaStr)
        Dim fcBachiller = New FieldContent("Bachiller", bachiller)
        Dim fcTituloTesis = New FieldContent("Titulo Tesis", tituloTesis)
        Dim fcAñoTesis = New FieldContent("Año Tesis", añoTesis)

        Dim valuesToFill = New Content()
        With valuesToFill
            .Fields.Add(fcNumero)
            .Fields.Add(fcFilial)
            .Fields.Add(fcFecha)
            .Fields.Add(fcBachiller)
            .Fields.Add(fcTituloTesis)
            .Fields.Add(fcAñoTesis)
        End With

        Using outputDocument = New TemplateProcessor("OutputTest.docx").SetRemoveContentControls(True)
            outputDocument.FillContent(valuesToFill)
            outputDocument.SaveChanges()
        End Using

    End Sub

End Module

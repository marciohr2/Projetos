'********************************************************************
'********************************************************************
'COMPONENTE CUSTOMIZADO, HERDA PROPRIEDADES DO TEXTBOX E ADICIONA
'NETFRAMEWORK 4 - PODE SER PORTADO PARA .NET MAIS RECENTES
'ALGUMAS FUNÇÕES
'CRIADO POR MÁRCIO RIBEIRO - 2013
'********************************************************************
'********************************************************************
Imports System.ComponentModel

Public Class MRSTextBox
    Inherits TextBox

    'Funcao da tecla enter
    Public Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef Iparam As String) As Integer

    Private mQuantidade_Decimais As Integer = 2
    Private mCor_Fundo_Ganhar_Foco As System.Drawing.Color = ColorTranslator.FromOle(RGB(251, 250, 221)) 'Amarelinho
    Private mCor_Fundo_Perder_Foco As System.Drawing.Color = ColorTranslator.FromOle(RGB(255, 255, 255)) 'Branco
    Private mEnter_Pula_Linhas As Boolean = True
    Private mSelecionar_Texto_ao_Ganhar_Foco As Boolean = True
    Private mMudar_Cor_Fonte_Numeros_Negativos As Boolean = True
    Private mCor_Fonte_Numeros_Negativos As System.Drawing.Color = Color.Red
    Private mCor_Fonte_Nao_Negativos As System.Drawing.Color = Color.Black

    Public Enum Tipos_Estilos
        Texto
        Numero_Inteiro
        Numero_Decimal
        Numero_Texto

        'Retornará sempre o indice, sendo
        'Texto = 0
        'Numero Inteiro = 1
        'Decimal = 2
        'Numeros Texto = 3

    End Enum

    Private mEstilo As Tipos_Estilos

    'PROPRIEDADES

    '****************************************************************************************************************************************
    'ESTILO
    '****************************************************************************************************************************************
    <CategoryAttribute("*Caixa de texto"), DisplayName("Estilo do texto"), DescriptionAttribute("Estilo de entrada do texto, Texto, Numeros Inteiros, Numeros decimais ou Numeros Texto")>
    Public Property Estilo() As Tipos_Estilos
        Get
            Return mEstilo
        End Get

        Set(ByVal value As Tipos_Estilos)
            mEstilo = value
            If mEstilo <> 0 Then 'Se for diferente de texto
                If IsNumeric(Me.Text) = False Then
                    Me.Text = "0"
                End If
            End If
        End Set
    End Property

    '****************************************************************************************************************************************
    'OPCAO DE MUDAR A COR DA FONTE PARA NUMEROS NEGATIVOS
    '****************************************************************************************************************************************
    <CategoryAttribute("*Caixa de texto"), DefaultValueAttribute(""), DisplayName("Cor da fonte negativos"), DescriptionAttribute("Cor da fonte ao receber numeros negativos (somente inteiros e decimais)")>
    Public Property Cor_Fonte_Numeros_Negativos() As System.Drawing.Color
        Get
            Return mCor_Fonte_Numeros_Negativos
        End Get
        Set(ByVal value As System.Drawing.Color)
            mCor_Fonte_Numeros_Negativos = value
        End Set
    End Property

    '****************************************************************************************************************************************
    'MUDAR COR DA FONTE PARA NUMEROS NEGATIVOS
    '****************************************************************************************************************************************
    <CategoryAttribute("*Caixa de texto"), DefaultValueAttribute("0"), DisplayName("Mudar cor da fonte negativos"), DescriptionAttribute("Mudar a cor da fonte ao receber numeros negativos (Somente inteiros e decimais)")>
    Public Property Mudar_Cor_Fonte_Numeros_Negativos() As Boolean
        Get
            Return mMudar_Cor_Fonte_Numeros_Negativos
        End Get
        Set(ByVal value As Boolean)
            mMudar_Cor_Fonte_Numeros_Negativos = value
        End Set
    End Property

    '****************************************************************************************************************************************
    'COR DA FONTE PARA NUMEROS NAO NEGATIVOS
    '****************************************************************************************************************************************
    <CategoryAttribute("*Caixa de texto"), DefaultValueAttribute(""), DisplayName("Cor da fonte não negativos"), DescriptionAttribute("Cor da fonte para não negativos, será usado no retorno da cor quando a caixa de texto não receber numeros negativos (somente inteiros e decimais)")>
    Public Property Cor_Fonte_Nao_Negativos() As System.Drawing.Color
        Get
            Return mCor_Fonte_Nao_Negativos
        End Get
        Set(ByVal value As System.Drawing.Color)
            mCor_Fonte_Nao_Negativos = value
        End Set
    End Property

    '****************************************************************************************************************************************
    'QUANTIDADE DE DECIMAIS
    '****************************************************************************************************************************************
    <CategoryAttribute("*Caixa de texto"), DefaultValueAttribute("0"), DisplayName("Quantidade de decimais"), DescriptionAttribute("Máximo de decimais suportar na caixa, desde que o estilo seja decimal")>
    Public Property Quantidade_Decimais() As Integer
        Get
            Return mQuantidade_Decimais
        End Get
        Set(ByVal value As Integer)
            mQuantidade_Decimais = value
        End Set
    End Property

    '****************************************************************************************************************************************
    'COR DO FUNDO AO GANHAR FOCO
    '****************************************************************************************************************************************
    <CategoryAttribute("*Caixa de texto"), DefaultValueAttribute(""), DisplayName("Cor de fundo ganhar foco"), DescriptionAttribute("Cor de fundo ao ganhar foco")>
    Public Property Cor_Fundo_Ganhar_Foco() As System.Drawing.Color
        Get
            Return mCor_Fundo_Ganhar_Foco
        End Get
        Set(ByVal value As System.Drawing.Color)
            mCor_Fundo_Ganhar_Foco = value
        End Set
    End Property

    '****************************************************************************************************************************************
    'COR DO FUNDO AO PERDER FOCO
    '****************************************************************************************************************************************
    <CategoryAttribute("*Caixa de texto"), DefaultValueAttribute(""), DisplayName("Cor de fundo perder foco"), DescriptionAttribute("Cor de fundo ao perder foco")>
    Public Property Cor_Fundo_Perder_Foco() As System.Drawing.Color
        Get
            Return mCor_Fundo_Perder_Foco
        End Get
        Set(ByVal value As System.Drawing.Color)
            mCor_Fundo_Perder_Foco = value
        End Set
    End Property

    '****************************************************************************************************************************************
    'ENTER PULA COMPONENTES
    '****************************************************************************************************************************************
    <CategoryAttribute("*Caixa de texto"), DefaultValueAttribute(True), DisplayName("Enter pula componentes"), DescriptionAttribute("Ao pressionar a tecla enter o sistema muda de foco para o próximo componente do form")>
    Public Property Enter_Pula_Linha() As Boolean
        Get
            Return mEnter_Pula_Linhas
        End Get
        Set(ByVal value As Boolean)
            mEnter_Pula_Linhas = value
        End Set
    End Property

    '****************************************************************************************************************************************
    'SELECIONAR TEXTO AO GANHAR FOCO
    '****************************************************************************************************************************************
    <CategoryAttribute("*Caixa de texto"), DefaultValueAttribute(True), DisplayName("Selecionar texto ao ganhar foco"), DescriptionAttribute("Ao focar na caixa de texto se houver conteúdo, será selecionado")>
    Public Property Selecionar_Texto_ao_Ganhar_Foco() As Boolean
        Get
            Return mSelecionar_Texto_ao_Ganhar_Foco
        End Get
        Set(ByVal value As Boolean)
            mSelecionar_Texto_ao_Ganhar_Foco = value
        End Set
    End Property


#Region "DESIGN DO COMPONENTE"

    Public Sub New(ByVal Container As System.ComponentModel.IContainer)
        MyClass.New()
        Container.Add(Me)
    End Sub

    Public Sub New()
        MyBase.New()
        InitializeComponent()
    End Sub


    'Substituições de componentes eliminam a lista de componentes.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If IsNothing(components) = False Then
                components.Dispose()
            End If
        End If

        MyBase.Dispose(disposing)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container()
    End Sub

#End Region

    '//////////////////////////////////////////////////////////////////////////////////////////////////////
    'PROPRIEDADES DO TEXTBOX!!!!
    '//////////////////////////////////////////////////////////////////////////////////////////////////////

    'Se o pião apagar tudo e estiver no estilo numerico, manda o caracter zero e seleciona tudo, pra evitar que a caixa esteja em branco
    Protected Overrides Sub OnTextChanged(ByVal e As System.EventArgs)

        If mEstilo = Tipos_Estilos.Numero_Inteiro Or mEstilo = Tipos_Estilos.Numero_Decimal Then
            If Me.Text = "" Then
                Me.Text = "0"
                Me.SelectAll()
            End If

            'Evita a colagem de dados texto, caso esteja no estilo numerico
            If Not IsNumeric(Me.Text) Then
                Me.Text = "0"
                Me.SelectAll()
            End If

        End If
        MyBase.OnTextChanged(e)
    End Sub

    Protected Overrides Sub OnEnter(ByVal e As System.EventArgs)
        Me.BackColor = mCor_Fundo_Ganhar_Foco

        If mSelecionar_Texto_ao_Ganhar_Foco = True Then
            Me.SelectAll()
        Else
            Me.SelectionStart = Len(Me.Text)
        End If

        MyBase.OnEnter(e)
    End Sub

    'Somente para tipo inteiro ou decimal
    Protected Overrides Sub OnClick(ByVal e As System.EventArgs)
        If mEstilo = Tipos_Estilos.Numero_Inteiro Or mEstilo = Tipos_Estilos.Numero_Decimal Or mEstilo = Tipos_Estilos.Numero_Texto Then
            If mSelecionar_Texto_ao_Ganhar_Foco = True Then
                Me.SelectAll()
            Else
                Me.SelectionStart = Len(Me.Text)
            End If
        End If

        MyBase.OnClick(e)
    End Sub

    Protected Overrides Sub OnLeave(ByVal e As System.EventArgs)
        Me.BackColor = mCor_Fundo_Perder_Foco

        'Verifica se está em branco a caixa, caso seja numerico
        If (mEstilo = Tipos_Estilos.Numero_Inteiro Or mEstilo = Tipos_Estilos.Numero_Decimal Or mEstilo = Tipos_Estilos.Numero_Texto) And (Not IsNumeric(Me.Text)) Then
            Me.Text = "0"
        End If

        'Se habilitado mudar a cor da fonte de numeros negativos, aplica
        If (mEstilo = Tipos_Estilos.Numero_Inteiro Or mEstilo = Tipos_Estilos.Numero_Decimal) And mMudar_Cor_Fonte_Numeros_Negativos = True Then
            If CDbl(Me.Text) < 0 Then
                Me.ForeColor = mCor_Fonte_Numeros_Negativos
            Else
                Me.ForeColor = mCor_Fonte_Nao_Negativos
            End If
        End If

        'Formata a caixa ao perder o foco, estilo inteiro
        If mEstilo = 1 Then
            Me.Text = CDbl(Me.Text)
        End If

        'Formata a caixa ao perder o foco, estilo decimal
        If mEstilo = 2 Then
            Me.Text = FormatNumber(Me.Text, mQuantidade_Decimais, TriState.True, TriState.False, TriState.False)
        End If


        'No estilo 3 não precisa formatar, deixa como foi digitado
        MyBase.OnLeave(e)

    End Sub

    Protected Overrides Sub OnKeyPress(ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Dim KeyAscii As Short = CShort(Asc(e.KeyChar))

        'Se permitir tudo
        If mEstilo = Tipos_Estilos.Texto Then
            If mEnter_Pula_Linhas = True Then
                'SendKeys.Send(vbTab)
                If KeyAscii = 13 Then
                    Call PostMessage(Me.Handle.ToInt32, &H100, &H9, 0) 'Pula linha com Enter
                End If
            End If
            Exit Sub
        End If

        'Permite outras teclas
        Select Case KeyAscii
            Case 8
                Exit Sub
            Case 13
                If mEnter_Pula_Linhas = True Then
                    'SendKeys.Send(vbTab)
                    Call PostMessage(Me.Handle.ToInt32, &H100, &H9, 0) 'Pula linha com Enter
                End If
                Exit Sub
            Case 32
                Exit Sub
        End Select



        'Verifica se digitou ponto
        If InStr(Me.Text, ".") > 0 Then
            Me.Text = Replace(Me.Text, ".", ",") 'Se houver ponto, substitui por virgula
            Me.SelectionStart = Len(Me.Text)
        End If


        '*********************************************************************************************************
        'Estilos para numeros
        '*********************************************************************************************************
        If mEstilo = Tipos_Estilos.Numero_Inteiro Then 'somente(numeros) inteiros
            KeyAscii = CShort(SoNumeros(KeyAscii))
            If KeyAscii = 0 Then
                e.Handled = True
            End If
        End If

        If mEstilo = Tipos_Estilos.Numero_Decimal Then 'decimais
            KeyAscii = CShort(SoNumeros(KeyAscii))
            If KeyAscii = 0 Then
                e.Handled = True
            End If
        End If

        If mEstilo = Tipos_Estilos.Numero_Texto Then 'Numeros texto
            KeyAscii = CShort(SoNumeros(KeyAscii))
            If KeyAscii = 0 Then
                e.Handled = True
            End If
        End If

        MyBase.OnKeyPress(e)
    End Sub


    Function SoNumeros(ByVal Keyascii As Short) As Short

        SoNumeros = 0


        '******************************************************************************
        'Verifica se já tem o negativo
        '******************************************************************************
        If Keyascii = 45 Then
            If InStr(Me.Text, "-") > 0 Then 'Se já houver negativo não deixa
                SoNumeros = 0
            Else
                Me.Text = "-" & Me.Text 'Se for negativar 
                Me.SelectionStart = Len(Me.Text)
                'SoNumeros = Keyascii   'Permite
                SoNumeros = 0
            End If
            Exit Function
        End If

        '******************************************************************************
        'Se pressionar o +, retira o menos se houver
        '******************************************************************************
        If Keyascii = 43 Then
            If InStr(Me.Text, "-") > 0 Then 'Se já houver ponto ou virgua, não deixa
                Dim Strvalor As String = Mid(Me.Text, 2, Len(Me.Text))
                Me.Text = Strvalor
                Me.SelectionStart = Len(Me.Text)
                SoNumeros = 0
            End If
            Exit Function
        End If


        '*********************************************************************************
        'ESTILO INTEIRO
        '*********************************************************************************
        If mEstilo = Tipos_Estilos.Numero_Inteiro Then
            If InStr("1234567890-", Chr(Keyascii)) = 0 Then
                SoNumeros = 0
            Else
                SoNumeros = Keyascii
            End If
            Exit Function
        End If

        '*********************************************************************************
        'ESTILO DECIMAL
        '*********************************************************************************
        If mEstilo = Tipos_Estilos.Numero_Decimal Then
            If InStr("1234567890,.-", Chr(Keyascii)) = 0 Then
                SoNumeros = 0
            Else

                '***************************************************************************
                'Se estiver tudo selecionado, manda ver, será apagado o conteudo
                '***************************************************************************
                If Me.SelectionLength = Len(Me.Text) Then
                    SoNumeros = Keyascii
                    Exit Function
                End If

                '******************************************************************************
                'Se a quantidade de decimais for maior que 0
                '******************************************************************************
                If mQuantidade_Decimais > 0 Then


                    '******************************************************************************
                    'Verifica se foi virgula ou ponto
                    '******************************************************************************
                    If Keyascii = 44 Or Keyascii = 46 Then
                        If InStr(Me.Text, ",") > 0 Then 'Se já houver ponto ou virgua, não deixa
                            SoNumeros = 0
                            Exit Function
                        Else
                            SoNumeros = 44 'Permite e envia a virgula, invés do ponto
                        End If
                    Else
                        SoNumeros = Keyascii 'Permite
                    End If



                    '******************************************************************************
                    'Se permitir decimais, verifica se já tem decimais
                    '******************************************************************************
                    If InStr(Me.Text, ",") > 0 Then 'Se já houver ponto ou virgua, verifica a quantidade de numeros após a virgula
                        If Permite_Mais_Decimais() = True Then
                            SoNumeros = Keyascii
                            Exit Function
                        Else
                            SoNumeros = 0
                            Exit Function
                        End If
                    End If

                Else
                    SoNumeros = 0
                End If
            End If
        End If

        '*********************************************************************************
        'ESTILO NUMERO TEXTO
        '*********************************************************************************
        If mEstilo = Tipos_Estilos.Texto Then
            If InStr("1234567890", Chr(Keyascii)) = 0 Then
                SoNumeros = 0
            Else
                SoNumeros = Keyascii
            End If
            Exit Function
        End If


    End Function


    'Verifica quantos numeros já foram digitados após a virgula
    Private Function Permite_Mais_Decimais() As Boolean
        Permite_Mais_Decimais = False

        Dim Posicao_virgula As Integer
        Dim Quantidade_Numeros_apos_virgula As Integer

        Posicao_virgula = InStr(Me.Text, ",")

        If Posicao_virgula > 0 Then
            Quantidade_Numeros_apos_virgula = Len(Mid(Me.Text, Posicao_virgula + 1, Len(Me.Text)))
            If Quantidade_Numeros_apos_virgula >= mQuantidade_Decimais Then
                Permite_Mais_Decimais = False
            Else
                Permite_Mais_Decimais = True
            End If
        End If
    End Function
End Class



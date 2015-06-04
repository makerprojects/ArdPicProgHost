Option Strict Off
Option Explicit On
Friend Class LED
	Inherits System.Windows.Forms.UserControl
	Public Event IntervalChange()
	Public Event StateChange()
    Public Event ColorChange()
	Public Event BlinkChange()
	' This control is an LED indicator control that can be placed on a form
	' to simulate an actual LED. The display indicator is painted as a round
	' image using direct drawing commands instead of the more common LED implementations
	' using bitmap images. The indicator image includes a 3D border and a reflection
	' glint to give the appearance of realism to the control. Through the use of the
	' graphic drawing algorithm used the LED can be set to almost any practical size
	' at design time from a few pixels across to few hundred pixels, although the
	' realism look of the indicator start to degrade after the control is sized over
	' about 50 or 60 pizels. (The ability to size the LED was a key motivator to design
	' and build this control as other sample controls from the WEB are fixed size and
	' many times too small for display screens with a high density pixels/inch setting.
	'
	' This control supports four property settings all of which are read/write at both
	' design time and run time. The properties are:
	'
	'   LED.State As Boolean                Controls the LED On/Off state. True = On while
	'                                       False = Off. (Note that this state is a higher
	'                                       state over the Blink states on.off that flashes
	'                                       the image.
	'
	'   LED.Blink As Boolean                This controls whether the LED will blink when it
	'                                       is in the On state. An LED can have its blink state
	'                                       set on even though the LED may be actually off via
	'                                       the State property. True = Blink enabled while
	'                                       False = no blink.
	'
	'   LED.Interval As Integer             This property specifies the blink rate interval for
	'                                       the LED in nominal units of milliseconds. Note that
	'                                       this property value is utilized directly by a standard
	'                                       timer and as such has similar granualrity and resolution
	'                                       behavior that a normal timer control will have. Practical
	'                                       values for the interval parameter are from 50 to 1000.
	'                                       At 1000 the LED blinks off for 1 second and then on for
	'                                       1 second.
	'
	'   LED.Color As LEDColorSelection      This property specifies the color for the LED. The values
	'                                       are specified by as Public enumeration value in the control.
	'                                       The selection values are:
	'
	'                                               0 - LED_Red
	'                                               1 - LED_Green
	'                                               2 - LED_Yellow
	'                                               3 - LED_Blue
	'
	'
    ' This code is based on the: 
    ' LED Display Indicator Control
    ' Written by: Michael Karas, Carousel Design Solutions (www.CarouselDesign.com)
    '
    ' Migration to VB 2008 by:
    ' Gregor Schlechtriem
    '
    ' You may incorporate all or parts of this program into your own project.
    ' Keep in mind that you got this code free and as such Carousel Design Solutions
    ' has no responsibility to you, your customers, or anyone else regarding the use
    ' of this program code. No support or warrenty of any kind can be provided. All
    ' liability falls upon the user of this code and thus you must judge the suitability
    ' of this code to your application and usage.
    '

	' specify the LED Colors that are supported
	Public Enum LEDColorSelection
		LED_Red
		LED_Green
		LED_Yellow
		LED_Blue
	End Enum
	
	'
	' declare the constants that represent the default values of the properties
	' in the user control
	'
	Private dState As Boolean ' default the on/off state of the LED
	Private dBlink As Boolean ' default control that indicates if an LED blinks
	Private dInterval As Short ' default blink interval property
	Private dColor As LEDColorSelection ' default color selection for the LED
	
	' local copies of the controls properties
	Private lState As Boolean ' the on/off state of the LED
	Private lBlink As Boolean ' control that indicates if an LED blinks
	Private lInterval As Short ' blink interval property
	Private lColor As LEDColorSelection ' color selection for the LED
	Private lOnOff As Boolean ' tracks the actual on/off state of LED
	
	'
	' property value routines for getting and setting the states of the user control
	'
	
	' property handlers for the State Property
	
	Public Property State() As Boolean
		Get
			
			State = lState
			
		End Get
		Set(ByVal Value As Boolean)
			
			lState = Value ' set the new state value
			lOnOff = Value
			If lState = False Then ' this logic sorts out the interaction
				BlinkTimer.Enabled = False ' to the On/Off state the blink state and the
			Else ' blink timer.
				If lBlink = True Then
                    BlinkTimer.Interval = lInterval
					BlinkTimer.Enabled = True
				End If
			End If
            RaiseEvent StateChange()
            Me.Invalidate(ClientRectangle)
		End Set
	End Property
	
	' property handlers for the Blink Property
	
	Public Property Blink() As Boolean
		Get
			
			Blink = lBlink
			
		End Get
		Set(ByVal Value As Boolean)
			
			lBlink = Value ' set the new blink state value
			RaiseEvent BlinkChange()
			
			If lState = True Then ' this logic sorts out the interaction
				lOnOff = True ' of the other control properties in relation
				If lBlink = True Then ' to the On/Off state the blink state and the
                    BlinkTimer.Interval = lInterval ' blink timer.
					BlinkTimer.Enabled = True
				Else
					BlinkTimer.Enabled = False
                End If
                Me.Invalidate(ClientRectangle)
            End If
			
		End Set
	End Property
	
	' property handlers for the Interval Property
	
	Public Property Interval() As Short
		Get
			
			Interval = lInterval
			
		End Get
		Set(ByVal Value As Short)
			
			lInterval = Value ' set the new interval value
			RaiseEvent IntervalChange()
			
			If BlinkTimer.Enabled = True Then ' Copy the interval to the blink
				BlinkTimer.Enabled = False ' timer if the timer is active for
                BlinkTimer.Interval = lInterval ' the disable / enable is to sync to
				BlinkTimer.Enabled = True ' interval right away.
			End If
			
		End Set
	End Property
	
	' property handlers for the Color Property
	
	Public Property Color() As LEDColorSelection
		Get
			
			Color = lColor
			
		End Get
		Set(ByVal Value As LEDColorSelection)
			
			lColor = Value ' select the new color setting
            RaiseEvent ColorChange() ' update the LED image
            Me.Invalidate(ClientRectangle)
        End Set
	End Property
	
	' Initialize the default states of the controls properties
	' this routine sets up some default value copies of the
	' properties so that the Write and Read properties methods
	' use the same default values as set here.
    Private Sub UserControl_InitProperties()

        dState = False
        Me.State = dState

        dBlink = False
        Me.Blink = dBlink

        dInterval = 0
        Me.Interval = dInterval

        dColor = LEDColorSelection.LED_Red
        Me.Color = dColor

    End Sub
	
    ' initialize default states of the control properties
    Private Sub UserControl_Initialize()

    End Sub
    Private Sub Circle(ByVal cp As PointF, ByVal r As Integer, ByVal myColor As Color, ByVal startAngle As Single, ByVal endAngle As Single)
        Dim myPen As New Pen(myColor, 2)
        Dim g As Graphics = Me.CreateGraphics
        Dim myRectangle As New Rectangle
        myRectangle.X = cp.X - r
        myRectangle.Y = cp.Y - r
        myRectangle.Width = 2 * r
        myRectangle.Height = 2 * r
        g.DrawArc(myPen, myRectangle, startAngle, endAngle - startAngle)
        g.Dispose()
    End Sub

	' routine to respond to the change of the LED blinking and toggle
	' its actual on/off state
	
	Private Sub BlinkTimer_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles BlinkTimer.Tick
		
		If lOnOff = False Then
			lOnOff = True
		Else
			lOnOff = False
        End If
        Me.Invalidate(ClientRectangle)
    End Sub

    ' display the LED image. this paints the image from the actual
    ' internal lOnOff boolean state so that this represents the
    ' actual displayed state whilest in the blinking state

    Private Sub LED_Paint(ByVal eventSender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        Dim XPos As Short ' X center position for the LED image
        Dim YPos As Short ' Y center position for the LED image
        Dim R As Integer
        Dim RBeg As Integer ' Beginning radius of the LED glint
        Dim RWid As Integer ' Width if the LED glint
        Dim RSiz As Integer ' nominal radius of the LED image
        Dim RColor As Color ' variable that holds the LED image color
        Dim GColor As Color ' variable that holds the LED  glint color
        Dim myPoint As New PointF ' center point of circle

        ' compute X & Y center across width of the control window
        XPos = (ClientRectangle.Width \ 2)
        YPos = (ClientRectangle.Height \ 2) ' compute the Y center across the height of the
        RSiz = ClientRectangle.Width ' compute the LED radius value from the smaller
        If RSiz > ClientRectangle.Height Then ' of the control width and height.
            RSiz = ClientRectangle.Height
        End If
        RSiz = (RSiz \ 2) - 2

        RBeg = LInt(RSiz, 5, 2, 19, 4) ' interpolate / extrapolate a proportional size
        RWid = LInt(RSiz, 5, 1, 19, 4) ' for the LED glint based on the image radius

        ' setup the color variables
        Select Case lColor

            Case LEDColorSelection.LED_Red
                If lOnOff = False Then
                    RColor = ColorTranslator.FromWin32(&H52)
                    GColor = SystemColors.ButtonFace
                Else
                    RColor = ColorTranslator.FromWin32(&HFF)
                    GColor = System.Drawing.Color.White
                End If

            Case LEDColorSelection.LED_Green
                If lOnOff = False Then
                    RColor = ColorTranslator.FromWin32(&H5200)
                    GColor = SystemColors.ButtonFace
                Else
                    RColor = ColorTranslator.FromWin32(&HFF00)
                    GColor = System.Drawing.Color.White
                End If

            Case LEDColorSelection.LED_Yellow
                If lOnOff = False Then
                    RColor = ColorTranslator.FromWin32(&H8484)
                    GColor = SystemColors.ButtonFace
                Else
                    RColor = ColorTranslator.FromWin32(&HFFFF)
                    GColor = System.Drawing.Color.White
                End If

            Case LEDColorSelection.LED_Blue
                If lOnOff = False Then
                    RColor = ColorTranslator.FromWin32(&H8B2000)
                    GColor = SystemColors.ButtonFace
                Else
                    RColor = ColorTranslator.FromWin32(&HFF1985)
                    GColor = System.Drawing.Color.White
                End If

            Case Else ' default to Red
                If lOnOff = False Then
                    RColor = ColorTranslator.FromWin32(&H5200)
                    GColor = SystemColors.ButtonFace
                Else
                    RColor = ColorTranslator.FromWin32(&HFF00)
                    GColor = System.Drawing.Color.White
                End If

        End Select

        myPoint.X = XPos
        myPoint.Y = YPos
        ' draw the 3D border of the LED
        Me.Circle(myPoint, RSiz + 1, SystemColors.ButtonHighlight, 0.0F, 360.0F) ' button highlight
        Me.Circle(myPoint, RSiz + 1, SystemColors.ButtonShadow, 135.0F, 325.0F) ' button shadow
        ' draw the filled circle of the LED at its current color
        For R = RSiz - 1 To 1 Step -1
            Me.Circle(myPoint, R, RColor, 0.0F, 360.0F)
        Next
        ' draw the LED glint over the top left side of the image
        For R = RSiz - RBeg To RSiz - RBeg - RWid Step -1
            Me.Circle(myPoint, R, GColor, 180.0F, 270.0F) ' button hilight
        Next R

    End Sub

    ' function to perform linear interpolation along a straight line segment
    ' this computes Y from an X value between two known reference points X1.Y1 and X2.Y2
    Private Function LInt(ByRef X As Integer, ByRef X1 As Integer, ByRef Y1 As Integer, ByRef X2 As Integer, ByRef Y2 As Integer) As Integer

        LInt = (((Y2 - Y1) * (X - X1)) / (X2 - X1)) + Y1

    End Function
End Class
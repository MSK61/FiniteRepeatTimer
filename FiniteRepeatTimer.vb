'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' program:      finite repeat timer

' file:         FiniteRepeatTimer.vb

' function:     methods of the FiniteRepeatTimer component

' description:  repeats a specific action for a limited number of times

' author:       Mohammed Safwat (MS)

' environment:  visual studio.net enterprise architect 2003,
'               windows xp professional()

' notes:        This is a private program.

' revisions:    1.00  6/1/2006 (MS)	first release

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Imports System.ComponentModel
Public Class FiniteRepeatTimer
    Inherits System.Windows.Forms.Timer
    ' This method raises the Tick event.
    Protected Overrides Sub OnTick(ByVal e As EventArgs)

        SyncLock Me

            itsCurrentTrigger += 1 ' Increment the trigger
            ' counter.
            MyBase.OnTick(e)

            If itsCurrentTrigger = Repetitions Then ' if this is
                ' the last trigger

                Enabled = False ' Stop the timer from invoking
                ' more triggers.

            End If

        End SyncLock

    End Sub ' end of method OnTick
    ' gets the current trigger of the timer
    <DesignerSerializationVisibility( _
        DesignerSerializationVisibility.Hidden), Browsable( _
        False)> Public ReadOnly Property CurrentTrigger( _
        ) As Short
        Get

            SyncLock Me

                CurrentTrigger = itsCurrentTrigger

            End SyncLock

        End Get
    End Property
    ' enables/disables the timer
    Public Overrides Property Enabled() As Boolean
        Get

            Enabled = MyBase.Enabled

        End Get
        Set(ByVal Value As Boolean)

            MyBase.Enabled = Value
            SyncLock Me

                itsCurrentTrigger = 0

            End SyncLock

        End Set
    End Property
    ' gets/sets the number of times the timer should be triggered
    Public Property Repetitions() As Short
        Get

            Repetitions = itsRepetitions

        End Get
        Set(ByVal Value As Short)

            itsRepetitions = Value
            SyncLock Me

                If Enabled Then

                    itsCurrentTrigger = 0

                End If

            End SyncLock

        End Set
    End Property
    ' the current trigger of the timer
    Private itsCurrentTrigger As Short = 0
    Private itsRepetitions As Short = 0 ' the number of
    ' repetitions the timer should be triggered
End Class ' end of class FiniteRepeatTimer

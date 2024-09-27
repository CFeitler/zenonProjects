using Scada.AddIn.Contracts;
using System;
using System.Reflection;
using Scada.AddIn.Contracts.Screen;
using static TouchEventsGesture.Constants;
using System.Security.Principal;

namespace TouchEventsGesture
{

  [AddInExtension("Touch Event Gesture", "Touch Event detection for picture gesture and element gesture recognition based on Microsoft interaction context.")]
  public class ProjectServiceExtension : IProjectServiceExtension
  {
    private IProject _currentProject;

    /// <summary>
    /// Screen in project which is set as test bed for API events.
    /// </summary>
    private const string SCREEN_NAME = "Api_PictureGesture";

    /// <summary>
    /// button in project which is set as test bed for API events. Must be in screen of name [SCREEN_NAME]
    /// </summary>
    private const string BUTTON_NAME = "Button_1";

    private const string SERVICE_STARTED_VARIABLE_NAME = "ServiceTouchGestureEventsStarted";
    private const double STARTED = 1.0;
    private const double STOPPED = 0.0;

    public void Start(IProject context, IBehavior behavior)
    {
      _currentProject = context;

      try
      {
        _currentProject.ScreenCollection.ElementGesture += ScreenCollection_ElementGesture;
        _currentProject.ScreenCollection.PictureGesture += ScreenCollection_PictureGesture;

        var startedVariable = _currentProject?.VariableCollection[SERVICE_STARTED_VARIABLE_NAME];
        if (startedVariable != null)
        {
          startedVariable.SetValue(0, STARTED);
          startedVariable.SetValue(STARTED, 0, 0, 0);
        }
      }
      catch (Exception e)
      {
        WriteToCel($"Exception during initialising of Gesture recognition service. {e}");
      }


      WriteToCel("Touch gesture recognition service started!");
    }

    public void Stop()
    {
      try
      {
        _currentProject.ScreenCollection.ElementGesture -= ScreenCollection_ElementGesture;

        _currentProject.ScreenCollection.PictureGesture -= ScreenCollection_PictureGesture;

        WriteToCel("Gesture recognition service stopped!");

        var startedVariable = _currentProject?.VariableCollection[SERVICE_STARTED_VARIABLE_NAME];
        if (startedVariable != null)
        {
          startedVariable.SetValue(0, STOPPED);
          startedVariable.SetValue(STOPPED,0,0,0);
        }
      }
      catch (Exception e)
      {
        WriteToCel($"Exception during stopping of service. {e}");
      }
    }

    private void ScreenCollection_PictureGesture(object sender, Scada.AddIn.Contracts.Screen.PictureGestureEventArgs e)
    {
      try
      {
        if (e.Screen.Name != SCREEN_NAME)
          return;

        // Picture Gesture Event Arguments contains an array: InteractionContext.
        // Elements of this array describe the type and parameters of the gesture
        var celMessage = "Picture gesture: " + GetCelMessage(e.InteractionContextOutput);
        WriteToCel(celMessage);
        SetMessageVariables(e.InteractionContextOutput);

      }
      catch (Exception exception)
      {
        WriteToCel($"Error while processing picture gesture! {exception}");
      }

    }
    private void ScreenCollection_ElementGesture(object sender, Scada.AddIn.Contracts.Screen.ElementGestureEventArgs e)
    {
      try
      {
        if (e.Element.Name != BUTTON_NAME)
          return;
        var screenNameOfElementName = e.Element.Parent.Parent.Name;
        if (screenNameOfElementName == SCREEN_NAME)
          return;
        // Element Gesture Event Arguments contains an array: InteractionContext.
        // Elements of this array describe the type and parameters of the gesture
        var celMessage = "Element gesture: " + GetCelMessage(e.InteractionContextOutput);
        WriteToCel(celMessage);
        SetMessageVariables(e.InteractionContextOutput);
      }
      catch (Exception exception)
      {
        WriteToCel($"Error while processing picture gesture! {exception}");
      }
    }

    private void SetMessageVariables(object[] ico)
    {
      string variableValue = "";
      // Picture Gesture Event Arguments contains an array: InteractionContext.
      // Elements of this array describe the type and parameters of the gesture
      //object[] ico = (object[])e.InteractionContextOutput;
      var indexInteractionId = 0;
      var indexInteractionFlag = 1;
      var indexPointerType = 2;
      var indexXPosition = 3;
      var indexYPosition = 4;
      var indexInteractionArguments = 5;

      // The first element indicates the type of gesture. 
      // see reference: https://learn.microsoft.com/en-us/windows/win32/api/interactioncontext/ne-interactioncontext-interaction_id
      var interactionId = (InteractionId)Convert.ToUInt32(ico[indexInteractionId]);
      if (interactionId == InteractionId.INTERACTION_ID_MANIPULATION)
        variableValue = "MANIPULATION";
      if (interactionId == InteractionId.INTERACTION_ID_TAP)
        variableValue = "TAP";
      if (interactionId == InteractionId.INTERACTION_ID_SECONDARY_TAP)
        variableValue = "SECONDARY TAP";
      if (interactionId == InteractionId.INTERACTION_ID_HOLD)
        variableValue = "HOLD";
      if (interactionId == InteractionId.INTERACTION_ID_DRAG)
        variableValue = "DRAG";
      if (interactionId == InteractionId.INTERACTION_ID_CROSS_SLIDE)
        variableValue = "CROSS SLIDE";

      WriteToStringVariable("GestureLastInteraction", variableValue);

      // The second element indicates the interaction flag. 
      // it tells if the interaction was started, ended or the inertia of the gesture
      variableValue = "";
      var interactionFlags = (InteractionFlags)Convert.ToUInt32(ico[indexInteractionFlag]);
      if (interactionFlags == 0)
        variableValue = "NONE";
      if ((interactionFlags & InteractionFlags.INTERACTION_FLAG_BEGIN) != 0)
        variableValue += "BEGIN,";
      if ((interactionFlags & InteractionFlags.INTERACTION_FLAG_END) != 0)
        variableValue += "END,";
      if ((interactionFlags & InteractionFlags.INTERACTION_FLAG_CANCEL) != 0)
        variableValue += "CANCEL,";
      if ((interactionFlags & InteractionFlags.INTERACTION_FLAG_INERTIA) != 0)
        variableValue += "INERTIA,";

      WriteToStringVariable("GestureLastInteractionFlag", variableValue);

      // The third element indicates the type of input element
      // zenon provides messages predominantly for TOUCH input
      var inputType = (PointerInputType)Convert.ToUInt32(ico[indexPointerType]);
      if (inputType == PointerInputType.POINTER)
        variableValue = "POINTER";
      if (inputType == PointerInputType.MOUSE)
        variableValue = "MOUSE";
      if (inputType == PointerInputType.PEN)
        variableValue = "PEN";
      if (inputType == PointerInputType.TOUCH)
        variableValue = "TOUCH";

      WriteToStringVariable("GestureLastPointerType", variableValue);

      // The fourth element (float) is the x position but in relative pixel position of the monitor 
      var xPosition = Math.Floor(Convert.ToDouble(ico[indexXPosition]));
      WriteToStringVariable("GestureLastXPosition", $"{xPosition}");

      // The fith   element (float) is the y position but in relative pixel position of the monitor 
      var yPosition = Math.Floor(Convert.ToDouble(ico[indexYPosition]));
      WriteToStringVariable("GestureLastYPosition", $"{yPosition}");

      // The type of the 6th element is dependant on the interaction type and hold the corresponding interaction arguments. 
      // In case of a TAP the element indicates the count. 
      if (interactionId == InteractionId.INTERACTION_ID_TAP)
      {
        var tapCount = Convert.ToUInt32(ico[indexInteractionArguments]);
        WriteToStringVariable("GestureLastTapCount", $"{tapCount}");
      }
      else
      {
        WriteToStringVariable("GestureLastTapCount", "");
      }

      // In case of a manipulation the arguments are an object itself
      if (interactionId == InteractionId.INTERACTION_ID_MANIPULATION)
      {
        object[] arguments = (object[])ico[indexInteractionArguments];

        var indexArgumentDelta = 0;
        var indexArgumentCumulative = 1;
        var indexArgumentVelocity = 2;
        var indexArgumentRailState = 3;

        var indexTranslationX = 0;
        var indexTranslationY = 1;
        var indexScale = 2;
        var indexExpansion = 3;
        var indexRotation = 4;

        object[] manipulationDelta = (object[])arguments[indexArgumentDelta];
        var deltaTranslationX = Convert.ToDouble(manipulationDelta[indexTranslationX]);
        WriteToStringVariable("GestureLastManipulationDeltaTranslationX", $"{deltaTranslationX}");
        var deltaTranslationY = Convert.ToDouble(manipulationDelta[indexTranslationY]);
        WriteToStringVariable("GestureLastManipulationDeltaTranslationY", $"{deltaTranslationY}");
        var deltaScale = Convert.ToDouble(manipulationDelta[indexScale]);
        WriteToStringVariable("GestureLastManipulationDeltaScale", $"{deltaScale}");
        var deltaExpansion = Convert.ToDouble(manipulationDelta[indexExpansion]);
        WriteToStringVariable("GestureLastManipulationDeltaExpansion", $"{deltaExpansion}");
        var deltaRotation = Convert.ToDouble(manipulationDelta[indexRotation]);
        WriteToStringVariable("GestureLastManipulationDeltaRotation", $"{deltaRotation}");

        object[] manipulationCumulative = (object[])arguments[indexArgumentCumulative];
        var cumulativeTranslationX = Convert.ToDouble(manipulationCumulative[indexTranslationX]);
        WriteToStringVariable("GestureLastManipulationCumulativeTranslationX", $"{cumulativeTranslationX}");
        var cumulativeTranslationY = Convert.ToDouble(manipulationCumulative[indexTranslationY]);
        WriteToStringVariable("GestureLastManipulationCumulativeTranslationY", $"{cumulativeTranslationY}");
        var cumulativeScale = Convert.ToDouble(manipulationCumulative[indexScale]);
        WriteToStringVariable("GestureLastManipulationCumulativeScale", $"{cumulativeScale}");
        var cumulativeExpansion = Convert.ToDouble(manipulationCumulative[indexExpansion]);
        WriteToStringVariable("GestureLastManipulationCumulativeExpansion", $"{cumulativeExpansion}");
        var cumulativeRotation = Convert.ToDouble(manipulationCumulative[indexRotation]);
        WriteToStringVariable("GestureLastManipulationCumulativeRotation", $"{cumulativeRotation}");

        var indexVelocityX = 0;
        var indexVelocityY = 1;
        var indexVelocityExpansion = 2;
        var indexVelocityAngular = 3;
        object[] manipulationVelocity = (object[])arguments[indexArgumentVelocity];
        var velocityX = Convert.ToDouble(manipulationVelocity[indexVelocityX]);
        WriteToStringVariable("GestureLastManipulationVelocityX", $"{velocityX}");
        var velocityY = Convert.ToDouble(manipulationVelocity[indexVelocityY]);
        WriteToStringVariable("GestureLastManipulationVelocityY", $"{velocityY}");
        var velocityExpansion = Convert.ToDouble(manipulationVelocity[indexVelocityExpansion]);
        WriteToStringVariable("GestureLastManipulationVelocityExpansion", $"{velocityExpansion}");
        var velocityAngular = Convert.ToDouble(manipulationVelocity[indexVelocityAngular]);
        WriteToStringVariable("GestureLastManipulationVelocityRotation", $"{velocityAngular}");

        ManipulationRailsState railsState = (ManipulationRailsState)Convert.ToUInt32(arguments[indexArgumentRailState]);
        WriteToStringVariable("GestureLastManipulationRailsState", $"{railsState}");

      }
      else
      {
        //clear the string values
        WriteToStringVariable("GestureLastManipulationDeltaTranslationX", "");
        WriteToStringVariable("GestureLastManipulationDeltaTranslationY", "");
        WriteToStringVariable("GestureLastManipulationDeltaScale", "");
        WriteToStringVariable("GestureLastManipulationDeltaExpansion", "");
        WriteToStringVariable("GestureLastManipulationDeltaRotation", "");

        WriteToStringVariable("GestureLastManipulationCumulativeTranslationX", "");
        WriteToStringVariable("GestureLastManipulationCumulativeTranslationY", "");
        WriteToStringVariable("GestureLastManipulationCumulativeScale", "");
        WriteToStringVariable("GestureLastManipulationCumulativeExpansion", "");
        WriteToStringVariable("GestureLastManipulationCumulativeRotation", "");

        WriteToStringVariable("GestureLastManipulationVelocityX", "");
        WriteToStringVariable("GestureLastManipulationVelocityY", "");
        WriteToStringVariable("GestureLastManipulationVelocityExpansion", "");
        WriteToStringVariable("GestureLastManipulationVelocityRotation", "");

        WriteToStringVariable("GestureLastManipulationRailsState", "");
      }

      if (interactionId == InteractionId.INTERACTION_ID_CROSS_SLIDE)
      {
        var crossSideFlags = (CrossSlideFlags)Convert.ToUInt32(ico[indexInteractionArguments]);
        if (crossSideFlags == CrossSlideFlags.NONE)
          variableValue = " NONE";
        if ((crossSideFlags & CrossSlideFlags.REARRANGE) != 0)
          variableValue = " REARRANGE,";
        if ((crossSideFlags & CrossSlideFlags.SELECT) != 0)
          variableValue = ", SELECT";
        if ((crossSideFlags & CrossSlideFlags.SPEED_BUMP) != 0)
          variableValue = ", SPEED BUMP";
        WriteToStringVariable("GestureLastCrossSlideFlags", variableValue);
      }
    }

    private static string GetCelMessage(object[] ico)
    {
      var message = "";

      // Picture Gesture Event Arguments contains an array: InteractionContext.
      // Elements of this array describe the type and parameters of the gesture
      var indexInteractionId = 0;
      var indexInteractionFlag = 1;
      var indexPointerType = 2;
      var indexXPosition = 3;
      var indexYPosition = 4;
      var indexInteractionArguments = 5;


      // The first element indicates the type of gesture. 
      // see reference: https://learn.microsoft.com/en-us/windows/win32/api/interactioncontext/ne-interactioncontext-interaction_id
      var interactionId = (InteractionId)Convert.ToUInt32(ico[indexInteractionId]);
      if (interactionId == InteractionId.INTERACTION_ID_MANIPULATION)
        message += "MANIPULATION";
      if (interactionId == InteractionId.INTERACTION_ID_TAP)
        message += "TAP";
      if (interactionId == InteractionId.INTERACTION_ID_SECONDARY_TAP)
        message += "SECONDARY TAP";
      if (interactionId == InteractionId.INTERACTION_ID_HOLD)
        message += "HOLD";
      if (interactionId == InteractionId.INTERACTION_ID_DRAG)
        message += "DRAG";
      if (interactionId == InteractionId.INTERACTION_ID_CROSS_SLIDE)
        message += "CROSS SLIDE";


      // The second element indicates the interaction flag. 
      // it tells if the interaction was started, ended or the inertia of the gesture
      message += " Flag(s): ";
      var interactionFlags = (InteractionFlags)Convert.ToUInt32(ico[indexInteractionFlag]);
      if (interactionFlags == 0)
        message += "NONE";
      if ((interactionFlags & InteractionFlags.INTERACTION_FLAG_BEGIN) != 0)
        message += "BEGIN,";
      if ((interactionFlags & InteractionFlags.INTERACTION_FLAG_END) != 0)
        message += "END,";
      if ((interactionFlags & InteractionFlags.INTERACTION_FLAG_CANCEL) != 0)
        message += "CANCEL,";
      if ((interactionFlags & InteractionFlags.INTERACTION_FLAG_INERTIA) != 0)
        message += "INERTIA,";

      // The third element indicates the type of input element
      // zenon provides messages predominantly for TOUCH input
      message += " using: ";
      var inputType = (PointerInputType)Convert.ToUInt32(ico[indexPointerType]);
      if (inputType == PointerInputType.POINTER)
        message += "POINTER";
      if (inputType == PointerInputType.MOUSE)
        message += "MOUSE";
      if (inputType == PointerInputType.PEN)
        message += "PEN";
      if (inputType == PointerInputType.TOUCH)
        message += "TOUCH";

      // The fourth element (float) is the x position but in relative pixel position of the monitor 
      var xPosition = Math.Floor(Convert.ToDouble(ico[indexXPosition]));
      message += $" X-Position: {xPosition}";
      // The fith   element (float) is the y position but in relative pixel position of the monitor 
      var yPosition = Math.Floor(Convert.ToDouble(ico[indexYPosition]));
      message += $" Y-Position: {yPosition}";

      // The type of the 6th element is dependant on the interaction type and hold the corresponding interaction arguments. 
      // In case of a TAP the element indicates the count. 
      if (interactionId == InteractionId.INTERACTION_ID_TAP)
      {
        var tapCount = Convert.ToUInt32(ico[indexInteractionArguments]);
        message += $" Count: {tapCount}";
      }

      // In case of a manipulation the arguments are an object itself
      if (interactionId == InteractionId.INTERACTION_ID_MANIPULATION)
      {
        object[] arguments = (object[])ico[indexInteractionArguments];
        if (arguments != null)
        {
          message += " to much args for CEL";
        }

        var indexArgumentDelta = 0;
        var indexArgumentCumulative = 1;
        var indexArgumentVelocity = 2;
        var indexArgumentRailState = 3;

        var indexTranslationX = 0;
        var indexTranslationY = 1;
        var indexScale = 2;
        var indexExpansion = 3;
        var indexRotation = 4;

        object[] manipulationDelta = (object[])arguments[indexArgumentDelta];
        var deltaTranslationX = Convert.ToDouble(manipulationDelta[indexTranslationX]);
        var deltaTranslationY = Convert.ToDouble(manipulationDelta[indexTranslationY]);
        var deltaScale = Convert.ToDouble(manipulationDelta[indexScale]);
        var deltaExpansion = Convert.ToDouble(manipulationDelta[indexExpansion]);
        var deltaRotation = Convert.ToDouble(manipulationDelta[indexRotation]);

        object[] manipulationCumulative = (object[])arguments[indexArgumentCumulative];
        var cumulativeTranslationX = Convert.ToDouble(manipulationCumulative[indexTranslationX]);
        var cumulativeTranslationY = Convert.ToDouble(manipulationCumulative[indexTranslationY]);
        var cumulativeScale = Convert.ToDouble(manipulationCumulative[indexScale]);
        var cumulativeExpansion = Convert.ToDouble(manipulationCumulative[indexExpansion]);
        var cumulativeRotation = Convert.ToDouble(manipulationCumulative[indexRotation]);

        var indexVelocityX = 0;
        var indexVelocityY = 1;
        var indexVelocityExpansion = 2;
        var indexVelocityAngular = 3;
        object[] manipulationVelocity = (object[])arguments[indexArgumentVelocity];
        var velocityX = Convert.ToDouble(manipulationVelocity[indexVelocityX]);
        var velocityY = Convert.ToDouble(manipulationVelocity[indexVelocityY]);
        var velocityExpansion = Convert.ToDouble(manipulationVelocity[indexVelocityExpansion]);
        var velocityAngular = Convert.ToDouble(manipulationVelocity[indexVelocityAngular]);

        ManipulationRailsState railsState = (ManipulationRailsState)Convert.ToUInt32(arguments[indexArgumentRailState]);
      }

      if (interactionId == InteractionId.INTERACTION_ID_CROSS_SLIDE)
      {
        var crossSideFlags = (CrossSlideFlags)Convert.ToUInt32(ico[indexInteractionArguments]);
        if (crossSideFlags == CrossSlideFlags.NONE)
          message += " NONE";
        if ((crossSideFlags & CrossSlideFlags.REARRANGE) != 0)
          message += " REARRANGE,";
        if ((crossSideFlags & CrossSlideFlags.SELECT) != 0)
          message += ", SELECT";
        if ((crossSideFlags & CrossSlideFlags.SPEED_BUMP) != 0)
          message += ", SPEED BUMP";
      }

      return message;
    }

    private void WriteToCel(string message)
    {
      var celVariable = _currentProject?.VariableCollection["CEL Message"];
      if (celVariable == null)
      {
        _currentProject?.ChronologicalEventList.AddEventEntry(
          $"{Assembly.GetExecutingAssembly().FullName}: Variable 'CEL Message' not found. This variable is needed to generate CEL entries!");
        return;
      }
      celVariable.SetValue(0, message);
    }

    private void WriteToStringVariable(string variableName, string message)
    {
      var stringVariable = _currentProject?.VariableCollection[variableName];
      if (stringVariable != null)
      {
        //stringVariable.SetValue(0, message);
        stringVariable.SetValue(message,0,0,0);
      }
    }
  }
}
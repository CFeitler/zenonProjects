using Scada.AddIn.Contracts;
using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.Remoting.Channels;
using System.Security.Cryptography;

namespace ScreenTouchEvents
{

  [AddInExtension("Screen Touch Events", "This service reacts on touch screen events on a screen called 'Api' and write the event type in the CEL.")]
  public class ProjectServiceExtension : IProjectServiceExtension
  {
    #region IProjectServiceExtension implementation

    private IProject _currentProject;

    /// <summary>
    /// Screen in project which is set as test bed for API events.
    /// </summary>
    private const string SCREEN_NAME = "Api";

    /// <summary>
    /// button in project which is set as test bed for API events. Must be in screen of name [SCREEN_NAME]
    /// </summary>
    private const string BUTTON_NAME = "Button_1";

    /// <summary>
    /// setting this to true activates also the events which fill up the CEL rapidly...
    /// </summary>
    private const bool USE_INTENSIVE_EVENTS = false;
    public void Start(IProject context, IBehavior behavior)
    {

      _currentProject = context;

      try
      {

        #region Raw Data Detection
        // These events will not be fired for windows 7 detection
        // These events work for mouse and touch pointer alike
        // These events only work, if the multi-touch events - raw data event routing in the
        // property Programming interface of the screen properties are selected or the selection is set to 'all events'.
        _currentProject.ScreenCollection.PointerActivate += ScreenCollection_PointerActivate;
        _currentProject.ScreenCollection.PointerDown += ScreenCollection_PointerDown;
        _currentProject.ScreenCollection.PointerEnter += ScreenCollection_PointerEnter;
        _currentProject.ScreenCollection.PointerLeave += ScreenCollection_PointerLeave;
        _currentProject.ScreenCollection.PointerUp += ScreenCollection_PointerUp;
        _currentProject.ScreenCollection.PointerCaptureChanged += ScreenCollection_PointerCaptureChanged;
        if (USE_INTENSIVE_EVENTS)
        {
          //Note: The event PointerUpdate creates a lot of events as all every change of the position of the mouse pointer is forwarded to the AddIn.
          _currentProject.ScreenCollection.PointerUpdate += ScreenCollection_PointerUpdate;
        }
        _currentProject.ScreenCollection.PointerWheel += ScreenCollection_PointerWheel;
        _currentProject.ScreenCollection.PointerHWheel += ScreenCollection_PointerHWheel;
        // NC stands for Non-Client. These events are hard to force. It work however occasionally if the touch point is on the system title.
        // see: https://learn.microsoft.com/en-us/windows/win32/inputmsg/wm-ncpointerdown
        _currentProject.ScreenCollection.NCPointerDown += ScreenCollection_NCPointerDown;
        _currentProject.ScreenCollection.NCPointerUp += ScreenCollection_NCPointerUp;
        _currentProject.ScreenCollection.NCPointerUpdate += ScreenCollection_NCPointerUpdate;
        #endregion

        #region Raw Data Detection Element (Mouse only)

        _currentProject.ScreenCollection.ElementLeftButtonDoubleClick += ScreenCollection_ElementLeftButtonDoubleClick;
        _currentProject.ScreenCollection.ElementLeftButtonDown += ScreenCollection_ElementLeftButtonDown;
        _currentProject.ScreenCollection.ElementLeftButtonUp += ScreenCollection_ElementLeftButtonUp;
        _currentProject.ScreenCollection.ElementRightButtonDoubleClick += ScreenCollection_ElementRightButtonDoubleClick;
        _currentProject.ScreenCollection.ElementRightButtonDown += ScreenCollection_ElementRightButtonDown;
        _currentProject.ScreenCollection.ElementRightButtonUp += ScreenCollection_ElementRightButtonUp;
        // The mouse over event cramps up the CEL als it delivers every mouse movement.
        if (USE_INTENSIVE_EVENTS)
        {
          _currentProject.ScreenCollection.ElementMouseOver += ScreenCollection_ElementMouseOver;
        }


        #endregion

        #region Windows 7 Events
        // I do not know how this 'touching' event is triggered. According to the zenon source code the touching event is only triggered with windows 7 configuration. 
        _currentProject.ScreenCollection.Touching += ScreenCollection_Touching;
        //works with windows 7 configuration:
        _currentProject.ScreenCollection.TouchManipulationStarted += ScreenCollection_TouchManipulationStarted;
        _currentProject.ScreenCollection.TouchManipulationCompleted += ScreenCollection_TouchManipulationCompleted;
        _currentProject.ScreenCollection.TouchManipulationDelta += ScreenCollection_TouchManipulationDelta;

        #endregion

        #region Gesture Recognition
        _currentProject.ScreenCollection.ElementGesture += ScreenCollection_ElementGesture;

        _currentProject.ScreenCollection.PictureGesture += ScreenCollection_PictureGesture;



        #endregion

        WriteToCel("Gesture recognition service started!");
      }
      catch (Exception e)
      {
        WriteToCel($"Exception during initialising of service. {e}");
      }

    }

    public void Stop()
    {

      try
      {
        #region Raw Data Detection
        _currentProject.ScreenCollection.PointerActivate -= ScreenCollection_PointerActivate;
        _currentProject.ScreenCollection.PointerDown -= ScreenCollection_PointerDown;
        _currentProject.ScreenCollection.PointerEnter -= ScreenCollection_PointerEnter;
        _currentProject.ScreenCollection.PointerLeave -= ScreenCollection_PointerLeave;
        _currentProject.ScreenCollection.PointerUp -= ScreenCollection_PointerUp;
        if (USE_INTENSIVE_EVENTS)
        {
          _currentProject.ScreenCollection.PointerUpdate -= ScreenCollection_PointerUpdate;
        }
        _currentProject.ScreenCollection.PointerWheel -= ScreenCollection_PointerWheel;
        _currentProject.ScreenCollection.PointerHWheel -= ScreenCollection_PointerHWheel;
        _currentProject.ScreenCollection.PointerCaptureChanged -= ScreenCollection_PointerCaptureChanged;
        _currentProject.ScreenCollection.NCPointerDown -= ScreenCollection_NCPointerDown;
        _currentProject.ScreenCollection.NCPointerUp -= ScreenCollection_NCPointerUp;
        _currentProject.ScreenCollection.NCPointerUpdate -= ScreenCollection_NCPointerUpdate;

        #endregion

        #region Raw Data Detection Element (Mouse only)

        _currentProject.ScreenCollection.ElementLeftButtonDoubleClick -= ScreenCollection_ElementLeftButtonDoubleClick;
        _currentProject.ScreenCollection.ElementLeftButtonDown -= ScreenCollection_ElementLeftButtonDown;
        _currentProject.ScreenCollection.ElementLeftButtonUp -= ScreenCollection_ElementLeftButtonUp;
        _currentProject.ScreenCollection.ElementRightButtonDoubleClick -= ScreenCollection_ElementRightButtonDoubleClick;
        _currentProject.ScreenCollection.ElementRightButtonDown -= ScreenCollection_ElementRightButtonDown;
        _currentProject.ScreenCollection.ElementRightButtonUp -= ScreenCollection_ElementRightButtonUp;
        if (USE_INTENSIVE_EVENTS)
        {
          _currentProject.ScreenCollection.ElementMouseOver -= ScreenCollection_ElementMouseOver;
        }

        #endregion

        #region Windows 7 Events
        // I do not know how this 'touching' event is triggered. According to the zenon source code the touching event is only triggered with windows 7 configuration. 
        _currentProject.ScreenCollection.Touching -= ScreenCollection_Touching;
        // works with windows 7 configuration:
        _currentProject.ScreenCollection.TouchManipulationStarted -= ScreenCollection_TouchManipulationStarted;
        _currentProject.ScreenCollection.TouchManipulationCompleted -= ScreenCollection_TouchManipulationCompleted;
        _currentProject.ScreenCollection.TouchManipulationDelta -= ScreenCollection_TouchManipulationDelta;
        #endregion

        #region Gesture recognition

        _currentProject.ScreenCollection.ElementGesture -= ScreenCollection_ElementGesture;

        _currentProject.ScreenCollection.PictureGesture -= ScreenCollection_PictureGesture;
        #endregion

        WriteToCel("Gesture recognition service stopped!");
      }
      catch (Exception e)
      {
        WriteToCel($"Exception during stopping of service. {e}");
      }

    }

    #region Constants and Flags from Interaction Context
    /// <summary>
    /// copied from https://learn.microsoft.com/en-us/windows/win32/api/interactioncontext/ne-interactioncontext-interaction_id
    /// </summary>
    enum InteractionId : uint
    {
      INTERACTION_ID_NONE = 0x00000000,
      INTERACTION_ID_MANIPULATION = 0x00000001,
      INTERACTION_ID_TAP = 0x00000002,
      INTERACTION_ID_SECONDARY_TAP = 0x00000003,
      INTERACTION_ID_HOLD = 0x00000004,
      INTERACTION_ID_DRAG = 0x00000005,
      INTERACTION_ID_CROSS_SLIDE = 0x00000006,
      INTERACTION_ID_MAX = 0xffffffff
    }

    /// <summary>
    /// copied from: https://learn.microsoft.com/en-us/windows/win32/api/interactioncontext/ne-interactioncontext-interaction_flags
    /// </summary>
    [Flags]
    enum InteractionFlags : uint
    {
      INTERACTION_FLAG_NONE = 0x00000000,
      INTERACTION_FLAG_BEGIN = 0x00000001,
      INTERACTION_FLAG_END = 0x00000002,
      INTERACTION_FLAG_CANCEL = 0x00000004,
      INTERACTION_FLAG_INERTIA = 0x00000008,
      INTERACTION_FLAG_MAX = 0xffffffff
    }

    enum PointerInputType
    {
      NONE = 0x00000000,
      POINTER = 0x00000001,
      TOUCH = 0x00000002,
      PEN = 0x00000003,
      MOUSE = 0x00000004
    }

    [Flags]
    enum ManipulationRailsState : uint
    {
      UNDECIDED = 0x00000000,
      FREE = 0x00000001,
      RAILED = 0x00000002
    }

    [Flags]
    internal enum CrossSideFlags : uint
    {
      NONE = 0x00000000,
      SELECT = 0x00000001,
      SPEED_BUMP = 0x00000002,
      REARRANGE = 0x00000004
    }

    /// <summary>
    /// copied from https://learn.microsoft.com/en-us/windows/win32/api/interactioncontext/ne-interactioncontext-interaction_configuration_flags
    /// </summary>
    enum INTERACTION_CONFIGURATION_FLAGS : uint
    {
      INTERACTION_CONFIGURATION_FLAG_NONE = 0x00000000,
      INTERACTION_CONFIGURATION_FLAG_MANIPULATION = 0x00000001,
      INTERACTION_CONFIGURATION_FLAG_MANIPULATION_TRANSLATION_X = 0x00000002,
      INTERACTION_CONFIGURATION_FLAG_MANIPULATION_TRANSLATION_Y = 0x00000004,
      INTERACTION_CONFIGURATION_FLAG_MANIPULATION_ROTATION = 0x00000008,
      INTERACTION_CONFIGURATION_FLAG_MANIPULATION_SCALING = 0x00000010,
      INTERACTION_CONFIGURATION_FLAG_MANIPULATION_TRANSLATION_INERTIA = 0x00000020,
      INTERACTION_CONFIGURATION_FLAG_MANIPULATION_ROTATION_INERTIA = 0x00000040,
      INTERACTION_CONFIGURATION_FLAG_MANIPULATION_SCALING_INERTIA = 0x00000080,
      INTERACTION_CONFIGURATION_FLAG_MANIPULATION_RAILS_X = 0x00000100,
      INTERACTION_CONFIGURATION_FLAG_MANIPULATION_RAILS_Y = 0x00000200,
      INTERACTION_CONFIGURATION_FLAG_MANIPULATION_EXACT = 0x00000400,
      INTERACTION_CONFIGURATION_FLAG_MANIPULATION_MULTIPLE_FINGER_PANNING = 0x00000800,
      INTERACTION_CONFIGURATION_FLAG_CROSS_SLIDE = 0x00000001,
      INTERACTION_CONFIGURATION_FLAG_CROSS_SLIDE_HORIZONTAL = 0x00000002,
      INTERACTION_CONFIGURATION_FLAG_CROSS_SLIDE_SELECT = 0x00000004,
      INTERACTION_CONFIGURATION_FLAG_CROSS_SLIDE_SPEED_BUMP = 0x00000008,
      INTERACTION_CONFIGURATION_FLAG_CROSS_SLIDE_REARRANGE = 0x00000010,
      INTERACTION_CONFIGURATION_FLAG_CROSS_SLIDE_EXACT = 0x00000020,
      INTERACTION_CONFIGURATION_FLAG_TAP = 0x00000001,
      INTERACTION_CONFIGURATION_FLAG_TAP_DOUBLE = 0x00000002,
      INTERACTION_CONFIGURATION_FLAG_TAP_MULTIPLE_FINGER,
      INTERACTION_CONFIGURATION_FLAG_SECONDARY_TAP = 0x00000001,
      INTERACTION_CONFIGURATION_FLAG_HOLD = 0x00000001,
      INTERACTION_CONFIGURATION_FLAG_HOLD_MOUSE = 0x00000002,
      INTERACTION_CONFIGURATION_FLAG_HOLD_MULTIPLE_FINGER,
      INTERACTION_CONFIGURATION_FLAG_DRAG = 0x00000001,
      INTERACTION_CONFIGURATION_FLAG_MAX = 0xffffffff
    }
    #endregion

    private void ScreenCollection_PictureGesture(object sender, Scada.AddIn.Contracts.Screen.PictureGestureEventArgs e)
    {
      try
      {
        var message = "Picture gesture: ";

        // Picture Gesture Event Arguments contains an array: InteractionContext.
        // Elements of this array describe the type and parameters of the gesture
        object[] ico = (object[])e.InteractionContextOutput;
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
          object[] manipulationVelocity = (object[])arguments[indexArgumentCumulative];
          var velocityX = Convert.ToDouble(manipulationVelocity[indexVelocityX]);
          var velocityY = Convert.ToDouble(manipulationVelocity[indexVelocityY]);
          var velocityExpansion = Convert.ToDouble(manipulationVelocity[indexVelocityExpansion]);
          var velocityAngular = Convert.ToDouble(manipulationVelocity[indexVelocityAngular]);

          ManipulationRailsState railsState = (ManipulationRailsState)Convert.ToUInt32(arguments[indexArgumentRailState]);

        }

        if (interactionId == InteractionId.INTERACTION_ID_CROSS_SLIDE)
        {
          var crossSideFlags = (CrossSideFlags)Convert.ToUInt32(ico[indexInteractionArguments]);
          if (crossSideFlags == CrossSideFlags.NONE)
            message += " NONE";
          if ((crossSideFlags & CrossSideFlags.REARRANGE) != 0)
            message += " REARRANGE,";
          if ((crossSideFlags & CrossSideFlags.SELECT) != 0)
            message += ", SELECT";
          if ((crossSideFlags & CrossSideFlags.SPEED_BUMP) != 0)
            message += ", SPEED BUMP";
        }


        WriteToCel(message);
      }
      catch (Exception exception)
      {
        WriteToCel($"Error while processing picture gesture! {exception}");
      }




    }

    private void ScreenCollection_ElementGesture(object sender, Scada.AddIn.Contracts.Screen.ElementGestureEventArgs e)
    {
      var screenNameOfElementName = e.Element.Parent.Parent.Name;
      if (screenNameOfElementName == SCREEN_NAME)
        return;
      WriteToCel($"Element Gesture detected on screen: {screenNameOfElementName} on element: {e.Element.Name} ");
    }

    #region Raw Data Detection
    private void ScreenCollection_PointerActivate(object sender, Scada.AddIn.Contracts.Screen.PointerActivateEventArgs e)
    {
      if (e.Screen.Name != SCREEN_NAME)
        return;
      WriteToCel($"from AddIn Service: {MethodBase.GetCurrentMethod().Name} recognized!");
    }


    private void ScreenCollection_NCPointerUpdate(object sender, Scada.AddIn.Contracts.Screen.NCPointerUpdateEventArgs e)
    {
      if (e.Screen.Name != SCREEN_NAME)
        return;
      WriteToCel($"from AddIn Service: {MethodBase.GetCurrentMethod().Name} recognized!");
    }

    private void ScreenCollection_NCPointerDown(object sender, Scada.AddIn.Contracts.Screen.NCPointerDownEventArgs e)
    {
      if (e.Screen.Name != SCREEN_NAME)
        return;
      WriteToCel($"from AddIn Service: {MethodBase.GetCurrentMethod().Name} recognized!");
    }

    private void ScreenCollection_PointerWheel(object sender, Scada.AddIn.Contracts.Screen.PointerWheelEventArgs e)
    {
      if (e.Screen.Name != SCREEN_NAME)
        return;
      WriteToCel($"from AddIn Service: {MethodBase.GetCurrentMethod().Name}: vertikal delta:{e.WheelDelta}");
    }

    private void ScreenCollection_PointerUp(object sender, Scada.AddIn.Contracts.Screen.PointerUpEventArgs e)
    {
      if (e.Screen.Name != SCREEN_NAME)
        return;
      WriteToCel($"from AddIn Service: {MethodBase.GetCurrentMethod().Name}: x:{e.PositionX} y:{e.PositionY}");
    }

    private void ScreenCollection_PointerEnter(object sender, Scada.AddIn.Contracts.Screen.PointerEnterEventArgs e)
    {
      if (e.Screen.Name != SCREEN_NAME)
        return;
      WriteToCel($"from AddIn Service: {MethodBase.GetCurrentMethod().Name} recognized!");
    }

    private void ScreenCollection_NCPointerUp(object sender, Scada.AddIn.Contracts.Screen.NCPointerUpEventArgs e)
    {
      if (e.Screen.Name != SCREEN_NAME)
        return;
      WriteToCel($"from AddIn Service: {MethodBase.GetCurrentMethod().Name} recognized!");
    }

    private void ScreenCollection_PointerHWheel(object sender, Scada.AddIn.Contracts.Screen.PointerHWheelEventArgs e)
    {
      if (e.Screen.Name != SCREEN_NAME)
        return;
      WriteToCel($"from AddIn Service: {MethodBase.GetCurrentMethod().Name} horizontal delta:{e.WheelDelta}");
    }

    private void ScreenCollection_PointerUpdate(object sender, Scada.AddIn.Contracts.Screen.PointerUpdateEventArgs e)
    {
      if (e.Screen.Name != SCREEN_NAME)
        return;
      WriteToCel($"from AddIn Service: {MethodBase.GetCurrentMethod().Name} recognized!");
    }

    private void ScreenCollection_PointerLeave(object sender, Scada.AddIn.Contracts.Screen.PointerLeaveEventArgs e)
    {
      if (e.Screen.Name != SCREEN_NAME)
        return;
      WriteToCel($"from AddIn Service: {MethodBase.GetCurrentMethod().Name} recognized!");
    }

    private void ScreenCollection_PointerDown(object sender, Scada.AddIn.Contracts.Screen.PointerDownEventArgs e)
    {
      if (e.Screen.Name != SCREEN_NAME)
        return;
      WriteToCel($"from AddIn Service: {MethodBase.GetCurrentMethod().Name} x:{e.PositionX} y:{e.PositionY} ");

    }

    private void ScreenCollection_PointerCaptureChanged(object sender,
      Scada.AddIn.Contracts.Screen.PointerCaptureChangedEventArgs e)
    {
      if (e.Screen.Name != SCREEN_NAME)
        return;
    }

    #endregion

    #region Raw Data Detection Element


    private void ScreenCollection_ElementMouseOver(object sender, Scada.AddIn.Contracts.Screen.ElementMouseOverEventArgs e)
    {
      if (e.Screen == null
          || e.Screen.Name != SCREEN_NAME
          || e.Element == null //in case the element is a static text this property could be null
          || e.Element.Name != BUTTON_NAME)
        return;
      WriteToCel($"from AddIn Service: {MethodBase.GetCurrentMethod().Name} on element: {e.Element.Name} recognized! x:{e.PositionX} y: {e.PositionY}");
    }

    private void ScreenCollection_ElementRightButtonUp(object sender, Scada.AddIn.Contracts.Screen.ElementRightButtonUpEventArgs e)
    {
      if (e.Screen == null
          || e.Screen.Name != SCREEN_NAME
          || e.Element == null
          || e.Element.Name != BUTTON_NAME)
        return;
      WriteToCel($"from AddIn Service: {MethodBase.GetCurrentMethod().Name} on element: {e.Element.Name} recognized! x:{e.PositionX} y: {e.PositionY}");
    }

    private void ScreenCollection_ElementRightButtonDown(object sender, Scada.AddIn.Contracts.Screen.ElementRightButtonDownEventArgs e)
    {
      if (e.Screen == null
          || e.Screen.Name != SCREEN_NAME
          || e.Element == null
          || e.Element.Name != BUTTON_NAME)
        return;
      WriteToCel($"from AddIn Service: {MethodBase.GetCurrentMethod().Name} on element: {e.Element.Name} recognized! x:{e.PositionX} y: {e.PositionY}");
    }

    private void ScreenCollection_ElementRightButtonDoubleClick(object sender, Scada.AddIn.Contracts.Screen.ElementRightButtonDoubleClickEventArgs e)
    {
      if (e.Screen == null
          || e.Screen.Name != SCREEN_NAME
          || e.Element == null
          || e.Element.Name != BUTTON_NAME)
        return;
      WriteToCel($"from AddIn Service: {MethodBase.GetCurrentMethod().Name} on element: {e.Element.Name} recognized! x:{e.PositionX} y: {e.PositionY}");
    }

    private void ScreenCollection_ElementLeftButtonUp(object sender, Scada.AddIn.Contracts.Screen.ElementLeftButtonUpEventArgs e)
    {
      if (e.Screen == null
          || e.Screen.Name != SCREEN_NAME
          || e.Element == null
          || e.Element.Name != BUTTON_NAME)
        return;
      WriteToCel($"from AddIn Service: {MethodBase.GetCurrentMethod().Name} on element: {e.Element.Name} recognized! x:{e.PositionX} y: {e.PositionY}");
    }

    private void ScreenCollection_ElementLeftButtonDown(object sender, Scada.AddIn.Contracts.Screen.ElementLeftButtonDownEventArgs e)
    {
      if (e.Screen == null
          || e.Screen.Name != SCREEN_NAME
          || e.Element == null
          || e.Element.Name != BUTTON_NAME)
        return;
      WriteToCel($"from AddIn Service: {MethodBase.GetCurrentMethod().Name} on element: {e.Element.Name} recognized! x:{e.PositionX} y: {e.PositionY}");
    }

    private void ScreenCollection_ElementLeftButtonDoubleClick(object sender, Scada.AddIn.Contracts.Screen.ElementLeftButtonDoubleClickEventArgs e)
    {
      if (e.Screen == null
          || e.Screen.Name != SCREEN_NAME
          || e.Element == null
          || e.Element.Name != BUTTON_NAME)
        return;
      WriteToCel($"from AddIn Service: {MethodBase.GetCurrentMethod().Name} on element: {e.Element.Name} recognized! x:{e.PositionX} y: {e.PositionY}");
    }


    #endregion

    #region Windows 7 Events
    private void ScreenCollection_TouchManipulationDelta(object sender, Scada.AddIn.Contracts.Screen.TouchManipulationDeltaEventArgs e)
    {
      WriteToCel($"Touch manipulation delta. x: {e.PositionX} y:{e.PositionY} fingers: {e.ContactCount}");
    }

    private void ScreenCollection_TouchManipulationCompleted(object sender, Scada.AddIn.Contracts.Screen.TouchManipulationCompletedEventArgs e)
    {
      WriteToCel($"Touch manipulation completed! x: {e.PositionX} y:{e.PositionY} fingers: {e.ContactCount}");
    }

    private void ScreenCollection_TouchManipulationStarted(object sender, Scada.AddIn.Contracts.Screen.TouchManipulationStartedEventArgs e)
    {
      WriteToCel($"Touch manipulation started x: {e.PositionX} y:{e.PositionY}");
    }

    private void ScreenCollection_Touching(object sender, Scada.AddIn.Contracts.Screen.TouchingEventArgs e)
    {
      WriteToCel($"Touching event. x-contacts:{string.Join(", ", e.XContacts)} y-contacts:{string.Join(", ", e.YContacts)}");
    }
    #endregion



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



    #endregion
  }
}
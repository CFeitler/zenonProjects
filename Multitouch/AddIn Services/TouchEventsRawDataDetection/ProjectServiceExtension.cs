using Scada.AddIn.Contracts;
using System;
using System.Reflection;

namespace TouchEventsRawDataDetection
{

  [AddInExtension("Touch Events Raw Data", "Touch Event detection for screens and elements. Raw Data is detected and printed on 'CEL Message' string variable.")]
  public class ProjectServiceExtension : IProjectServiceExtension
  {
    private IProject _currentProject;

    /// <summary>
    /// Screen in project which is set as test bed for API events.
    /// </summary>
    private const string SCREEN_NAME = "Api_RawData";

    /// <summary>
    /// button in project which is set as test bed for API events. Must be in screen of name [SCREEN_NAME]
    /// </summary>
    private const string BUTTON_NAME = "Button_1";

    /// <summary>
    /// setting this to true activates also the events which fill up the CEL rapidly...
    /// </summary>
    private const bool USE_INTENSIVE_EVENTS = false;

    private const string SERVICE_STARTED_VARIABLE_NAME = "ServiceTouchRawDataEventsStarted";
    private const double STARTED = 1.0;
    private const double STOPPED = 0.0;

    public void Start(IProject context, IBehavior behavior)
    {
      _currentProject = context;
      try
      {


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

        var startedVariable = _currentProject?.VariableCollection[SERVICE_STARTED_VARIABLE_NAME];
        if (startedVariable != null)
        {
          startedVariable.SetValue(0, STARTED);
        }

        WriteToCel("Touch recognition raw data service started!");
      }
      catch (Exception e)
      {
        WriteToCel($"Exception during initialising of Touch Raw Data service. {e}");
      }

    }

    public void Stop()
    {
      try
      {
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

        var startedVariable = _currentProject?.VariableCollection[SERVICE_STARTED_VARIABLE_NAME];
        if (startedVariable != null)
        {
          startedVariable.SetValue(0, STOPPED);
        }

        WriteToCel("Touch recognition raw data service stopped!");
      }
      catch (Exception e)
      {
        WriteToCel($"Exception during stopping of raw data service. {e}");
      }
    }

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

  }
}
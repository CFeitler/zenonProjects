using Scada.AddIn.Contracts;
using System;
using System.Reflection;

namespace TouchEventsWindows7
{

  [AddInExtension("Windows 7 touch events", "This service extension reacts on window 7 touch events. These events are only fired when the touch recognition is set to Windows 7")]
  public class ProjectServiceExtension : IProjectServiceExtension
  {

    private IProject _currentProject;

    /// <summary>
    /// Screen name in project which is set as test bed for Windows 7 API events.
    /// The screen name is not used by windows 7 events. The constant is kept for
    /// sake of continuity.
    /// </summary>
    private const string SCREEN_NAME = "Api_Windows7";

    private const string SERVICE_STARTED_VARIABLE_NAME = "ServiceTouchWindows7Started";
    private const double STARTED = 1.0;
    private const double STOPPED = 0.0;

    public void Start(IProject context, IBehavior behavior)
    {
      try
      {

        // I do not know how this 'touching' event is triggered. According to the zenon source code the touching event is only triggered with windows 7 configuration. 
        _currentProject.ScreenCollection.Touching += ScreenCollection_Touching;
        //works with windows 7 configuration:
        _currentProject.ScreenCollection.TouchManipulationStarted += ScreenCollection_TouchManipulationStarted;
        _currentProject.ScreenCollection.TouchManipulationCompleted += ScreenCollection_TouchManipulationCompleted;
        _currentProject.ScreenCollection.TouchManipulationDelta += ScreenCollection_TouchManipulationDelta;
        var startedVariable = _currentProject?.VariableCollection[SERVICE_STARTED_VARIABLE_NAME];
        if (startedVariable != null)
        {
          startedVariable.SetValue(0, STARTED);
        }
      }
      catch (Exception e)
      {
        WriteToCel($"Exception during initializing of service. {e}");
      }

    }

    public void Stop()
    {
      try
      {
        // I do not know how this 'touching' event is triggered. According to the zenon source code the touching event is only triggered with windows 7 configuration. 
        _currentProject.ScreenCollection.Touching -= ScreenCollection_Touching;
        // works with windows 7 configuration:
        _currentProject.ScreenCollection.TouchManipulationStarted -= ScreenCollection_TouchManipulationStarted;
        _currentProject.ScreenCollection.TouchManipulationCompleted -= ScreenCollection_TouchManipulationCompleted;
        _currentProject.ScreenCollection.TouchManipulationDelta -= ScreenCollection_TouchManipulationDelta;
        var startedVariable = _currentProject?.VariableCollection[SERVICE_STARTED_VARIABLE_NAME];
        if (startedVariable != null)
        {
          startedVariable.SetValue(0, STOPPED);
        }
      }
      catch (Exception e)
      {
        WriteToCel($"Exception during stopping of service. {e}");
      }

    }

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
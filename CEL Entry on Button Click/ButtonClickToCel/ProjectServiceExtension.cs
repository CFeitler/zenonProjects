using Scada.AddIn.Contracts;
using System;
using System.Collections.Generic;
using System.Runtime.Remoting.Contexts;

namespace ButtonClickToCel
{
  /// <summary>
  /// Description of Project Service Extension.
  /// </summary>
  [AddInExtension("ButtonClickToCel", "This Service Engine AddIn Serivce reacts on button clicks and writes a cel entry.")]
  public class ProjectServiceExtension : IProjectServiceExtension
  {
    #region IProjectServiceExtension implementation

    private IProject _project;

    //This is a list of buttons which the service reacts to. Other Buttons
    //are ignored. This list and the following flags could also be in a
    //configuration file.
    private List<string> _elementNamesToReact = new List<string>()
      { "Button_2_1_1", "Button_2_1", "Button_2_1_2", "Button_2_1_2_1" };

    private bool _printElementId = true;
    private bool _printUserName = true;
    private bool _printHelpText = true;

    private string _currentUser = "you";

    public void Start(IProject context, IBehavior behavior)
    {
      _project = context;
      context.ScreenCollection.ElementLeftButtonUp += ScreenCollection_ElementLeftButtonUp;
      if (_printUserName)
      {
        _currentUser = context.UserAdministration.LoggedInUserName;
        _project.UserAdministration.UserChanged += UserAdministration_UserChanged;
      }

    }

    private void UserAdministration_UserChanged(object sender, Scada.AddIn.Contracts.UserAdministration.UserChangedEventArgs e)
    {
      _currentUser = _project.UserAdministration.LoggedInUserName;
    }

    private void ScreenCollection_ElementLeftButtonUp(object sender, Scada.AddIn.Contracts.Screen.ElementLeftButtonUpEventArgs e)
    {
      if (e?.Element?.Name != null
        && _elementNamesToReact.Contains(e.Element.Name))
      {
        if (_printElementId)
        {
          _project.ChronologicalEventList.AddEventEntry($"{_currentUser} clicked on: {e.Element.Name}");
        }

        if (_printHelpText)
        {
          _project.ChronologicalEventList.AddEventEntry(e.Element.HelpFile);
        }
      }
    }

    public void Stop()
    {
      _project.ScreenCollection.ElementLeftButtonUp -= ScreenCollection_ElementLeftButtonUp; 
      _project.UserAdministration.UserChanged -= UserAdministration_UserChanged;
    }

    #endregion
  }
}
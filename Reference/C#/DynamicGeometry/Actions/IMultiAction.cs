using System;
using System.Collections.Generic;
using System.Text;

namespace GuiLabs.Utils.Actions
{
	public interface IMultiAction : IAction, IList<IAction>
	{
	}
}

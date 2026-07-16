using System;
using System.Collections.Generic;
using System.Text;

namespace DynamicGeometry
{
    public class CompileResult
    {
        public Func<double, double> Function;
        public Func<double> Expression { get; set; }
        public readonly List<IFigure> Dependencies = new List<IFigure>();
        public readonly List<CompileError> Errors = new List<CompileError>();

        public bool IsSuccess
        {
            get
            {
                return Errors.IsEmpty() && (Expression != null || Function != null);
            }
        }

        public void AddError(string error)
        {
            Errors.Add(new CompileError()
            {
                Text = error
            });
        }

        public void AddBindError(string figureName)
        {
            AddError(string.Format("Could not find figure with name '{0}'", figureName));
        }

        public void AddPropertyNotFoundError(IFigure figure, string propertyName)
        {
            AddError(string.Format("Could not find property '{0}' on figure '{1}'", propertyName, figure.Name));
        }

        public void AddMethodNotFoundError(string functionName)
        {
            AddError(string.Format("Could not find method '{0}'", functionName));
        }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            foreach (var error in Errors)
            {
                if (sb.Length > 0)
                {
                    sb.AppendLine();
                }
                sb.Append(error.Text);
            }
            return sb.ToString();
        }

        public void AddDependencyCycleError(string figureName)
        {
            AddError(string.Format("Using figure '{0}' will create a cycle. Circular dependencies are not allowed.", figureName));
        }

        public void AddUnknownIdentifierError(string text)
        {
            AddError(string.Format("Unknown identifier: '{0}'", text));
        }

        public void AddFigureIsNotAPointError(string longestPrefix)
        {
            AddError(string.Format("Figure '{0}' is not a point."));
        }

        public void AddIncorrectNumberOfArgumentsError(System.Reflection.MethodInfo method, int actualNumberOfArguments)
        {
            AddError(string.Format("Function '{0}' expects {1} arguments, and it was passed {2}",
                method.Name, method.GetParameters().Length, actualNumberOfArguments));
        }
    }
}

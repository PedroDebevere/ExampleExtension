using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

using EnvDTE;

namespace TestExtension
{
	internal static class DteExtensions
	{
		public static readonly string CSharpProjectGuid = "{FAE04EC0-301F-11D3-BF4B-00C04F79EFBC}";

		public static IEnumerable<Project> OnlyCSharp(this IEnumerable<Project> projects)
		{
			return projects.WithKind(CSharpProjectGuid);
		}

		public static IEnumerable<Project> WithKind(this IEnumerable<Project> projects, string kindGuid)
		{
			return projects.Where(p => String.Equals(p.Kind, kindGuid, StringComparison.OrdinalIgnoreCase));
		}

		public static IEnumerable<Project> GetAllProjects(this EnvDTE.Solution solution)
		{
			return solution.Projects
				.OfType<Project>()
				.SelectMany(GetChildProjects)
				.Union(solution.Projects.Cast<Project>())
				.Where(p => { try { return !string.IsNullOrEmpty(p.FullName); } catch { return false; } });
		}

		public static IEnumerable<Project> GetChildProjects(this Project parent)
		{
			try
			{
				const string vsProjectKindSolutionFolder = "{66A26720-8FB5-11D2-AA7E-00C04F688DDE}";
				if (parent.Kind != vsProjectKindSolutionFolder && parent.Collection == null)  // Unloaded
					return Enumerable.Empty<Project>();

				if (!string.IsNullOrEmpty(parent.FullName))
					return new[] { parent };
			}
			catch (COMException)
			{
				return Enumerable.Empty<Project>();
			}

			return parent.ProjectItems
				.OfType<ProjectItem>()
				.Where(p => p.SubProject != null)
				.SelectMany(p => GetChildProjects(p.SubProject));
		}
	}
}

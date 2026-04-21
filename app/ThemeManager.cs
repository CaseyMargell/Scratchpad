using System.Windows;

namespace Scratchpad;

public static class ThemeManager
{
    public static void Apply(bool light)
    {
        var dicts = Application.Current.Resources.MergedDictionaries;

        // Always re-add the base theme (Themes.xaml) as the first merged dict,
        // then optionally layer Light overrides on top.
        dicts.Clear();
        dicts.Add(new ResourceDictionary { Source = new Uri("pack://application:,,,/Themes.xaml") });
        if (light)
        {
            dicts.Add(new ResourceDictionary { Source = new Uri("pack://application:,,,/LightTheme.xaml") });
        }
    }
}

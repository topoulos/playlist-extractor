namespace PlaylistExtractor;

public static class Utility
{
    public static double ConvertEmuToPixel(long emus)
    {
        return emus / 9525.0;
    }
    
    public static string TimeStringFromSeconds(double d)
    {
        var timeSpan = TimeSpan.FromSeconds(d);

        string hoursString = (int)timeSpan.TotalHours > 0 ? ((int)timeSpan.TotalHours).ToString("00") + ":" : "";
        string minutesString = timeSpan.Minutes.ToString("00");
        string secondsString = timeSpan.Seconds.ToString("00");

        return $"{hoursString}{minutesString}:{secondsString}";
    }

}
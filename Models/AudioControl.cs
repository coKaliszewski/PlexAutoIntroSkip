using System.Collections.Generic;
using System.Linq;
using System.Threading;
using CSCore.CoreAudioAPI;

namespace PlexAutoIntroSkip.Models
{
    public class AudioControlSettings
    {
        public double AverageVolume => AudioLogs?.Average() ?? 0;
        public List<double> AudioLogs { get; set; }
        public int VolumeSliderPosition { get; set; }
        public double VolumeTarget { get; set; }
        public AudioSessionControl2 AudioProcess { get; set; }
        public int AudioProcessID { get; set; }
        public double VolumeTolerance { get; set; }
        public int VolumeGatherTimeInSeconds { get; set; }
    }
}

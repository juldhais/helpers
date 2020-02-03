using System;

namespace Helpers
{
    public class Speedometer
    {
        public DateTime StartTime { get; private set; }
        public DateTime EndTime { get; private set; }

        public void Start()
        {
            StartTime = DateTime.Now;
            EndTime = default;
        }

        public void Stop()
        {
            EndTime = DateTime.Now;
        }

        public TimeSpan GetElapsedTime()
        {
            if (StartTime == default) return new TimeSpan();

            if (EndTime == default)
                return DateTime.Now - StartTime;

            return EndTime - StartTime;
        }

        public double GetSpeed(int traveled)
        {
            var totalSeconds = GetElapsedTime().TotalSeconds;
            if (totalSeconds == 0) return 0;

            return traveled / totalSeconds;
        }

        public TimeSpan GetRemainingTime(int traveled, int distance)
        {
            var speed = GetSpeed(traveled);
            if (speed == 0) return new TimeSpan();

            var remainingSeconds = (distance - traveled) / GetSpeed(traveled);
            return TimeSpan.FromSeconds(remainingSeconds);
        }
    }
}

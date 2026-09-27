using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DynamicGeometry;

/// <summary>
/// Controls the behavior when another request is scheduled within the time span
/// after a previous request.
/// </summary>
public struct ThrottleOptions : IEquatable<ThrottleOptions>
{
    /// <summary>
    /// If a new request is scheduled within the timespan after a previous request,
    /// it replaces the previous request.
    /// If true, the last request runs.
    /// If false, the first request runs.
    /// </summary>
    public bool ReplaceExistingRequests;

    /// <summary>
    /// Resets the timer to a full interval when a new request is made.
    /// If false, the next action is taken when the current time interval completes.
    /// If true, the current in-progress timer is discarded and a new one is started instead.
    /// Starts a new timer at scheduling time, if needed.
    /// </summary>
    /// <remarks><see cref="StartNewTimeIntervalIfNewRequests"/> is ignored when this field is set.</remarks>
    public bool ResetTimeIntervalOnNewRequests;

    /// <summary>
    /// Upon time interval completion, if new requests have been made within the past interval,
    /// a new timer is started again and no action is taken. This ensures that there is
    /// always a delay of at least the specified time interval before the request is executed.
    /// If false, the request may run sooner than the current time interval elapses, as a
    /// result of the previously running timer expiring.
    /// Starts a new timer on wakeup, if needed.
    /// </summary>
    public bool StartNewTimeIntervalIfNewRequests;

    /// <summary>
    /// Lock around delegate execution so only one delegate can run at a time for the given subject.
    /// </summary>
    public bool Isolate;

    /// <summary>
    /// Indicates that work should be run immediately if there were no recent requests,
    /// but should not occur more than once in the given time interval.
    /// </summary>
    public bool RunImmediatelyIfFree;

    public static ThrottleOptions RunOnceImmediatelyIfFree = new ThrottleOptions
    {
        RunImmediatelyIfFree = true,
        ReplaceExistingRequests = true,
        StartNewTimeIntervalIfNewRequests = false,
        Isolate = true
    };

    public static ThrottleOptions PreferRunningAtLeastOncePerInterval = new ThrottleOptions
    {
        ReplaceExistingRequests = true,
        StartNewTimeIntervalIfNewRequests = false,
        Isolate = true
    };

    public static ThrottleOptions PostponeUntilNoNewRequests = new ThrottleOptions
    {
        ReplaceExistingRequests = true,
        StartNewTimeIntervalIfNewRequests = true,
        Isolate = true
    };

    public static ThrottleOptions FullIntervalAfterLastRequest = new ThrottleOptions
    {
        ReplaceExistingRequests = true,
        ResetTimeIntervalOnNewRequests = true,
        Isolate = true
    };

    public bool Equals(ThrottleOptions other)
    {
        return (ReplaceExistingRequests, ResetTimeIntervalOnNewRequests, StartNewTimeIntervalIfNewRequests, Isolate, RunImmediatelyIfFree) ==
            (other.ReplaceExistingRequests, other.ResetTimeIntervalOnNewRequests, other.StartNewTimeIntervalIfNewRequests, other.Isolate, other.RunImmediatelyIfFree);
    }

    public override int GetHashCode()
    {
        return (ReplaceExistingRequests, ResetTimeIntervalOnNewRequests, StartNewTimeIntervalIfNewRequests, Isolate, RunImmediatelyIfFree).GetHashCode();
    }

    public override string ToString()
    {
        return $"ReplaceExistingRequests: {ReplaceExistingRequests} ResetTimeIntervalOnNewRequests: {ResetTimeIntervalOnNewRequests} StartNewTimeIntervalIfNewRequests: {StartNewTimeIntervalIfNewRequests} Isolate: {Isolate} RunImmediately: {RunImmediatelyIfFree}";
    }

    public override bool Equals(object obj)
    {
        if (obj is ThrottleOptions other)
        {
            return Equals(other);
        }

        return false;
    }

    public static bool operator ==(ThrottleOptions left, ThrottleOptions right) => left.Equals(right);
    public static bool operator !=(ThrottleOptions left, ThrottleOptions right) => !left.Equals(right);
}

/// <summary>
/// Callbacks run on the thread pool: anything that touches controls posts to the UI thread.
/// </summary>
public class Throttle
{
    private static readonly Dictionary<object, ScheduledWork> scheduledActionsForSubject = new Dictionary<object, ScheduledWork>();

    private class ScheduledWork
    {
        public object Subject;
        public object Callback;
        public Task DelayTask;
        public DateTime LastScheduledTimeUtc;
        public TimeSpan RequestedDelay;
        public ThrottleOptions Options;
    }

    /// <summary>
    /// Schedules an action to execute for a subject after a delay of the specified length.
    /// If another action is scheduled while waiting for the delay, various debouncing
    /// behaviors can be chosen to limit how often requests run.
    /// </summary>
    public static void Schedule<T>(T subject, Action<T> action, TimeSpan delay, ThrottleOptions options = default)
    {
        if (options == default)
        {
            options = ThrottleOptions.PostponeUntilNoNewRequests;
        }

        lock (scheduledActionsForSubject)
        {
            var now = DateTime.UtcNow;
            if (scheduledActionsForSubject.TryGetValue(subject, out ScheduledWork existingScheduledWork))
            {
                if (options.ReplaceExistingRequests || existingScheduledWork.Callback == null)
                {
                    existingScheduledWork.Callback = action;
                }

                if (options.ResetTimeIntervalOnNewRequests)
                {
                    StartNewTimer<T>(existingScheduledWork);

                    // no need to mark the time of last request as we have an entire timer
                    // dedicated to our request and when it expires we know it's us
                    return;
                }

                existingScheduledWork.LastScheduledTimeUtc = now;
            }
            else
            {
                var work = new ScheduledWork()
                {
                    Subject = subject,
                    Callback = action,
                    LastScheduledTimeUtc = DateTime.MinValue,
                    Options = options,
                    RequestedDelay = delay,
                };
                scheduledActionsForSubject[subject] = work;

                StartNewTimer<T>(work, options.RunImmediatelyIfFree);
            }
        }
    }

    private static void StartNewTimer<T>(ScheduledWork work, bool runImmediately = false)
    {
        var delayTask = Task.Delay(work.RequestedDelay);
        work.DelayTask = delayTask;

        if (runImmediately)
        {
            Task.Run(() => OnWakeUp(delayTask, (T)work.Subject, remove: false));
        }

        _ = delayTask.ContinueWith(t => OnWakeUp(t, (T)work.Subject), TaskScheduler.Default);
    }

    private static void OnWakeUp<T>(Task delayTask, T subject, bool remove = true, bool isFlushing = false)
    {
        Action<T> callbackToRun = null;
        bool isolate = true;

        lock (scheduledActionsForSubject)
        {
            var now = DateTime.UtcNow;
            if (scheduledActionsForSubject.TryGetValue(subject, out var work))
            {
                if (delayTask != work.DelayTask)
                {
                    // the timespan this continuation resumes after is not the last one
                    // so just drop it on the floor
                    return;
                }

                if (!isFlushing &&
                    work.Options.StartNewTimeIntervalIfNewRequests &&
                    !work.Options.ResetTimeIntervalOnNewRequests)
                {
                    if (now - work.LastScheduledTimeUtc < work.RequestedDelay)
                    {
                        StartNewTimer<T>(work);
                        return;
                    }
                }

                callbackToRun = work.Callback as Action<T>;
                isolate = work.Options.Isolate;

                if (remove)
                {
                    scheduledActionsForSubject.Remove(subject);
                }
                else
                {
                    work.Callback = null;
                }
            }
        }

        ExecuteCallback(subject, callbackToRun, isolate);
    }

    private static void ExecuteCallback<T>(T subject, Action<T> callbackToRun, bool isolate)
    {
        if (callbackToRun == null)
        {
            return;
        }

        // now execute the potentially long-running action outside the scheduler lock
        if (isolate)
        {
            // only allow one at a time per subject if Isolate is true (default)
            lock (subject)
            {
                try
                {
                    callbackToRun(subject);
                }
                catch
                {
                }
            }
        }
        else
        {
            try
            {
                callbackToRun(subject);
            }
            catch
            {
            }
        }
    }

    public static bool IsScheduled<T>(T subject)
    {
        lock (scheduledActionsForSubject)
        {
            return scheduledActionsForSubject.ContainsKey(subject);
        }
    }

    /// <summary>
    /// Aborts potentially scheduled tasks for subject
    /// </summary>
    /// <returns>true if there were scheduled tasks, false otherwise</returns>
    public static bool Abort<T>(T subject)
    {
        lock (scheduledActionsForSubject)
        {
            if (scheduledActionsForSubject.TryGetValue(subject, out var work))
            {
                scheduledActionsForSubject.Remove(subject);
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Forces the scheduled work to run immediately as if the timer just expired
    /// </summary>
    public static bool Flush<T>(T subject)
    {
        ScheduledWork work;

        lock (scheduledActionsForSubject)
        {
            scheduledActionsForSubject.TryGetValue(subject, out work);
        }

        if (work != null)
        {
            OnWakeUp(work.DelayTask, subject, isFlushing: true);
            return true;
        }

        return false;
    }
}

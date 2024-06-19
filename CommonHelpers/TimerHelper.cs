using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace TMHelper.Common
{
    public delegate void TimerElapsedFunction();
    public delegate void TimerTimeOutFunction();

    public class TimerHelper
    {
        #region Properties
        private bool blnHasTimeOut { get; set; }
        private double dblTimerInterval { get; set; }
        private double dblMaxTimerDuration { get; set; }
        private Timer timer { get; set; }
        private TimerElapsedFunction timerElapsedFunction { get; set; }
        private TimerTimeOutFunction timerTimeOutFunction { get; set; }
        #endregion

        #region Constructor
        public TimerHelper(double timerInterval, double maxTimerDuration, TimerElapsedFunction timerElapsedHandler, TimerTimeOutFunction timerTimeOutHandler)
        {
            this.blnHasTimeOut = true;
            this.dblTimerInterval = timerInterval;
            this.dblMaxTimerDuration = maxTimerDuration;
            this.timerElapsedFunction = timerElapsedHandler;
            this.timerTimeOutFunction = timerTimeOutHandler;
        }

        public TimerHelper(double maxTimerDuration, TimerElapsedFunction timerElapsedHandler, TimerTimeOutFunction timerTimeOutHandler)
        {
            this.blnHasTimeOut = true;
            this.dblTimerInterval = 1000;
            this.dblMaxTimerDuration = maxTimerDuration;
            this.timerElapsedFunction = timerElapsedHandler;
            this.timerTimeOutFunction = timerTimeOutHandler;
        }

        public TimerHelper(double maxTimerDuration, TimerTimeOutFunction timerTimeOutHandler)
        {
            this.blnHasTimeOut = true;
            this.dblTimerInterval = 1000;
            this.dblMaxTimerDuration = maxTimerDuration;
            this.timerTimeOutFunction = timerTimeOutHandler;
        }

        public TimerHelper(double timerInterval, double maxTimerDuration, TimerTimeOutFunction timerTimeOutHandler)
        {
            this.blnHasTimeOut = true;
            this.dblTimerInterval = timerInterval;
            this.dblMaxTimerDuration = maxTimerDuration;
            this.timerTimeOutFunction = timerTimeOutHandler;
        }

        public TimerHelper(double timerInterval, TimerElapsedFunction timerElapsedHandler)
        {
            this.blnHasTimeOut = false;
            this.dblTimerInterval = timerInterval;
            this.timerElapsedFunction = timerElapsedHandler;
        }
        #endregion

        #region Public Method
        public void Start()
        {
            try
            {
                this.timer = new Timer(this.dblTimerInterval);
                this.timer.Elapsed += new ElapsedEventHandler(DefaultTimerElapsedFunction);
                this.timer.Enabled = true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void Stop()
        {
            try
            {
                if (this.timer != null)
                {
                    this.timer.Enabled = false;
                    this.timer.Dispose();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region Private Method
        private void DefaultTimerElapsedFunction(object sender, ElapsedEventArgs e)
        {
            try
            {
                if (this.blnHasTimeOut)
                {
                    this.timer.Enabled = false;

                    this.dblMaxTimerDuration = this.dblMaxTimerDuration - this.dblTimerInterval;
                    if (this.dblMaxTimerDuration < 0)
                    {
                        this.timer.Elapsed -= DefaultTimerElapsedFunction;
                        this.timer.Enabled = false;
                        this.timer.Dispose();
                        this.timerTimeOutFunction?.Invoke();
                    }
                    else
                    {
                        this.timerElapsedFunction?.Invoke();
                        this.timer.Enabled = true;
                    }
                }
                else
                {
                    this.timer.Enabled = false;
                    this.timerElapsedFunction?.Invoke();
                    this.timer.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion
    }
}

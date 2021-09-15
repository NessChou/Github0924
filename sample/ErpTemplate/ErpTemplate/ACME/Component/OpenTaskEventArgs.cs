using System;
using System.Collections.Generic;
using System.Text;

namespace ACME
{
    /// <summary>
    /// class for open task event arguments
    /// </summary>
    public class OpenTaskEventArgs : EventArgs
    {
        #region Data members

        int _taskID = 0;
        
        #endregion

        #region Constructor

        /// <summary>
        /// constructor 
        /// </summary>
        /// <param name="taskID">task id</param>
        public OpenTaskEventArgs(int taskID)
            : base()
        {
            _taskID = taskID;
        }

        
        #endregion

        #region Property

        /// <summary>
        /// get task id
        /// </summary>
        public int TaskID
        {
            get { return _taskID; }
        }
        
        #endregion
    }
}

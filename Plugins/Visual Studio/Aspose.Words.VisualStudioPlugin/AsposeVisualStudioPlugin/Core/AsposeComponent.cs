// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AsposeVisualStudioPluginWords.Core
{
    /// <summary>
    /// 
    /// </summary>
    public class AsposeComponent
    {
        private string _name;

        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }
        private bool _selected;

        public bool Selected
        {
            get { return _selected; }
            set { _selected = value; }
        }
        private string _downloadUrl;

        public string DownloadUrl
        {
            get { return _downloadUrl; }
            set { _downloadUrl = value; }
        }
        private string _downloadFileName;

        public string DownloadFileName
        {
            get { return _downloadFileName; }
            set { _downloadFileName = value; }
        }
        private bool _downloaded;

        public bool Downloaded
        {
            get { return _downloaded; }
            set { _downloaded = value; }
        }
        private string _currentVersion;

        public string CurrentVersion
        {
            get { return _currentVersion; }
            set { _currentVersion = value; }
        }
        private string _latestVersion;

        public string LatestVersion
        {
            get { return _latestVersion; }
            set { _latestVersion = value; }
        }
        private bool _latestRelease;

        public bool LatestRelease
        {
            get { return _latestRelease; }
            set { _latestRelease = value; }
        }
        private string _changeLog;

        public string ChangeLog
        {
            get { return _changeLog; }
            set { _changeLog = value; }
        }
        private string _remoteExamplesRepository;

        public string RemoteExamplesRepository
        {
            get { return _remoteExamplesRepository; }
            set { _remoteExamplesRepository = value; }
        }

   
        /// <summary>
        /// 
        /// </summary>
        public void AsposeJavaComponent()
        {
            _selected = false;
            _downloaded = false;
            _latestRelease = false;
        }
        /**
         * @return the _name
         */
        public string get_name()
        {
            return _name;
        }
       
        /**
         * @param _name the _name to set
         */
        public void set_name(string _name)
        {
            this._name = _name;
        }

        /**
         * @return the _selected
         */
        public bool is_selected()
        {
            return _selected;
        }

        /**
         * @param _selected the _selected to set
         */
        public void set_selected(bool _selected)
        {
            this._selected = _selected;
        }


        /**
         * @return the _downloaded
         */
        public bool is_downloaded()
        {
            return _downloaded;
        }

        /**
         * @param _downloaded the _downloaded to set
         */
        public void set_downloaded(bool _downloaded)
        {
            this._downloaded = _downloaded;
        }

        /**
         * @return the _currentVersion
         */
        public string get_currentVersion()
        {
            return _currentVersion;
        }

        /**
         * @param _currentVersion the _currentVersion to set
         */
        public void set_currentVersion(string _currentVersion)
        {
            this._currentVersion = _currentVersion;
        }

        /**
         * @return the _downloadUrl
         */
        public string get_downloadUrl()
        {
            return _downloadUrl;
        }

        /**
         * @param _downloadUrl the _downloadUrl to set
         */
        public void set_downloadUrl(string _downloadUrl)
        {
            this._downloadUrl = _downloadUrl;
        }

        /**
         * @return the _latestVersion
         */
        public string get_latestVersion()
        {
            return _latestVersion;
        }

        /**
         * @param _latestVersion the _latestVersion to set
         */
        public void set_latestVersion(string _latestVersion)
        {
            this._latestVersion = _latestVersion;
        }

        /**
         * @return the _latestRelease
         */
        public bool is_latestRelease()
        {
            return _latestRelease;
        }

        /**
         * @param _latestRelease the _latestRelease to set
         */
        public void set_latestRelease(bool _latestRelease)
        {
            this._latestRelease = _latestRelease;
        }

        /**
         * @return the _changeLog
         */
        public string get_changeLog()
        {
            return _changeLog;
        }

        /**
         * @param _changeLog the _changeLog to set
         */
        public void set_changeLog(string _changeLog)
        {
            this._changeLog = _changeLog;
        }

        /**
         * @return the _downloadFileName
         */
        public string get_downloadFileName()
        {
            return _downloadFileName;
        }

        /**
         * @param _downloadFileName the _downloadFileName to set
         */
        public void set_downloadFileName(string _downloadFileName)
        {
            this._downloadFileName = _downloadFileName;
        }
        public string get_remoteExamplesRepository()
        {
            return _remoteExamplesRepository;
        }

        /**
         * 
         * @param _remoteExamplesRepository
         */
        public void set_remoteExamplesRepository(string _remoteExamplesRepository)
        {
            this._remoteExamplesRepository = _remoteExamplesRepository;
        }

    }
}

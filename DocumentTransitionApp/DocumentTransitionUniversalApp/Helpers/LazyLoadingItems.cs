using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Threading.Tasks;
using Windows.UI.Xaml.Controls;

namespace DocumentTransitionUniversalApp.Helpers
{
    public class LazyLoadingItems<Elements> : INotifyPropertyChanged
    {
        #region Fields

        public ObservableCollection<Elements> Items { get; private set; }

        public bool IsPullRefresh
        {
            get
            {
                return _isPullRefresh;
            }

            set
            {
                _isPullRefresh = value;
                OnPropertyChanged(nameof(IsPullRefresh));
            }
        }

        ObservableCollection<Elements> _sourceItems { get; set; }
        private int _pageSize = 20;
        private int _lastIndex = 0;
        bool _isPullRefresh = false;
        ScrollViewer _scrollViewer;

        #endregion

        #region Constructors

        public LazyLoadingItems(ObservableCollection<Elements> items, ScrollViewer scrollViewer)
        {
            Items = new ObservableCollection<Elements>();
            _sourceItems = items;
            _scrollViewer = scrollViewer;
            SetScrollViewer();
            FillItems();
        }

        #endregion

        #region Public methods

        public event PropertyChangedEventHandler PropertyChanged;

        public void OnPropertyChanged(string name)
        {
            this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        #endregion

        #region Private methods

        private void SetScrollViewer()
        {
            _scrollViewer.ViewChanged += _scrollViewer_ViewChanged;
        }

        private void FillItems()
        {
            if (_sourceItems.Count < _pageSize)
                _pageSize = _sourceItems.Count;

            for (int i = 0; i < _pageSize; i++, _lastIndex++)
            {
                Items.Add(_sourceItems[i]);
            }
        }

        private async void _scrollViewer_ViewChanged(object sender, ScrollViewerViewChangedEventArgs e)
        {
            var sv = sender as ScrollViewer;

            if (!e.IsIntermediate)
            {
                Items.Clear();
                if (sv.ScrollableHeight - sv.VerticalOffset < 200.0)
                {
                    IsPullRefresh = true;
                    await Task.Delay(50);
                    int maxScroll = _lastIndex + _pageSize > _sourceItems.Count - 1 ? _sourceItems.Count : _lastIndex + _pageSize;
                    for (int i = _lastIndex; i < maxScroll; i++, _lastIndex++)
                    {
                        Items.Add(_sourceItems[i]);
                    }

                    sv.ChangeView(null, sv.VerticalOffset, null);
                }

                IsPullRefresh = false;
            }
        }

        #endregion
    }
}

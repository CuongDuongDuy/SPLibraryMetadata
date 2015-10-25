using System;
using System.Collections.ObjectModel;

namespace SPLibraryMetadata
{
    public class PagingViewModel
    {
        #region Properties

        private PagingNavigationMove CurrentMove { get; set; }

        public PagingStatus Status { get; set; }

        public int TotalRows { get; set; }

        private int rowsPerPage;

        public int RowsPerPage
        {
            get { return rowsPerPage; }
            set
            {
                rowsPerPage = value;
                UpdateNavigation(PagingNavigationMove.Reset);
            }
        }

        public ObservableCollection<int> RowsPerPageOptions
        {
            get
            {
                return new ObservableCollection<int>
                {
                    30,
                    50,
                    100,
                    200,
                };
            }
        }

        public int CurrentPage { get; set; }

        public int TotalPages { get; set; }

        public bool HasPrevious { get; set; }

        public bool HasNext { get; set; }

        public Action<PagingNavigationMove> MovePreviousCommand { get; set; }
        public Action<PagingNavigationMove> MoveNextCommand { get; set; }
        public Action<PagingNavigationMove> RowsPerPageCommand { get; set; }

        private Action<PagingIntegrationMetadata> OutCallbackAction { get; set; }

        #endregion

        #region Constructor and initialization

        public PagingViewModel(int rowsPerPage, Action<PagingIntegrationMetadata> outCallback, Action<PagingIntegrationCallbacks> integration)
        {
            MovePreviousCommand = (parm) => UpdateNavigation(PagingNavigationMove.Previous);
            MoveNextCommand = (parm) => UpdateNavigation(PagingNavigationMove.Next);
            RowsPerPageCommand = (parm) => UpdateNavigation(PagingNavigationMove.Reset);

            var integrationParameter = new PagingIntegrationCallbacks
            {
                CallIn = UpdatePagingMetadata,
                CallOut = outCallback
            };

            OutCallbackAction = outCallback; 

            integration.Invoke(integrationParameter);

            #region Default Values

            TotalRows = 0;
            CurrentPage = 1;
            TotalPages = 0;

            this.rowsPerPage = rowsPerPage;
            Status = PagingStatus.Initializing;
            CurrentMove = PagingNavigationMove.Reset;

            ExcuteOutCallback();

            #endregion
        }

        #endregion

        #region Update Navigation buttons, UI

        private void UpdateNavigation(PagingNavigationMove navigationMove)
        {
            CurrentMove = navigationMove;
            UpdateCurrentPage(navigationMove);
            UpdateNavigationUi();
            ExcuteOutCallback();
        }

        private void UpdateCurrentPage(PagingNavigationMove navigationMove)
        {
            switch (navigationMove)
            {
                case PagingNavigationMove.Previous:
                    if (CurrentPage > 1)
                    {
                        CurrentPage--;
                    }
                    break;
                case PagingNavigationMove.Next:
                    if (CurrentPage < TotalPages)
                    {
                        CurrentPage++;
                    }
                    break;
                case PagingNavigationMove.Reset:
                    CurrentPage = 1;
                    break;
            }
            Status = Status != PagingStatus.Initializing ? PagingStatus.Loading : PagingStatus.Initializing;
        }

        private void UpdateNavigationUi()
        {
            TotalPages = TotalRows % RowsPerPage == 0 ? TotalRows / RowsPerPage : TotalRows / RowsPerPage + 1;
            HasNext = CurrentPage != TotalPages;
        }

        private void ExcuteOutCallback()
        {
            if (OutCallbackAction == null) return;
            OutCallbackAction.Invoke(GetPagingMetadata());
        }


        #endregion

        #region Get/Update metadata

        private PagingIntegrationMetadata GetPagingMetadata()
        {
            var result = new PagingIntegrationMetadata
            {
                Status = Status,
                CurrentPage = CurrentPage,
                ItemsPerPage = RowsPerPage,
                Move = CurrentMove
            };
            return result;
        }

        public void UpdatePagingMetadata(PagingIntegrationMetadata metadata)
        {
            if (metadata == null) return;
            Status = metadata.Status;
            if (metadata.Move != PagingNavigationMove.Reset) return;
            TotalRows = metadata.TotalRows;
            CurrentPage = 1;
            UpdateNavigationUi();
        }

        #endregion

    }

}

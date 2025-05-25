import React, { useEffect, useState } from 'react';
import {
  ColumnActionsMode,
  ConstrainMode,
  ContextualMenu,
  DetailsList,
  DetailsListLayoutMode,
  DirectionalHint,
  IColumn,
  IContextualMenuProps,
  IDetailsList,
  Stack,
} from '@fluentui/react';

import { useSelection } from '../../hooks/useSelection';
import { useLoadStore } from '../../hooks/useLoadStore';
import { useAppDispatch, useAppSelector } from '../../store/hooks';

import { CommandBar } from './CommandBar';
import { GridFooter } from './GridFooter';
import { GridCell } from './GridCell';

import {
  clearChangedRecords,
  clearChangedRecordsAfterRefresh,
  deleteRecords,
  readdChangedRecordsAfterDelete,
  saveRecords,
} from '../../store/features/RecordSlice';
import { setLoading } from '../../store/features/LoadingSlice';
import {
  addNewRow,
  readdNewRowsAfterDelete,
  removeNewRows,
  setRows,
} from '../../store/features/DatasetSlice';

import { Row, Column, mapDataSetColumns,
  mapDataSetRows, getColumnsTotalWidth } from '../../mappers/dataSetMapper';
import { _onRenderDetailsHeader } from '../../styles/RenderStyles';
import { buttonStyles } from '../../styles/ButtonStyles';
import { gridStyles } from '../../styles/DetailsListStyles';
import { IDataSetProps } from '../AppWrapper';
import { getContainerHeight } from '../../utils/commonUtils';
import { clearInvalidFields } from '../../store/features/ErrorSlice';
import { Entity, ErrorDetails } from '../../services/DataverseService';

const ASC_SORT = 0;
const DESC_SORT = 1;

export const EditableGrid = ({ _service, _setContainerHeight,
  dataset, isControlDisabled, width }: IDataSetProps) => {
  const { selection, selectedRecordIds } = useSelection();

  const rows: Row[] = useAppSelector(state => state.dataset.rows);
  const newRows: Row[] = useAppSelector(state => state.dataset.newRows);
  const columns = mapDataSetColumns(dataset, _service);
  const isPendingDelete = useAppSelector(state => state.record.isPendingDelete);
  const isPendingLoad = useAppSelector(state => state.dataset.isPending);
  const [sortMenuProps, setSortMenuProps] = useState<IContextualMenuProps | undefined>(undefined);

  const dispatch = useAppDispatch();

  const detailsListRef = React.createRef<IDetailsList>();

  const resetScroll = () => {
    detailsListRef.current?.scrollToIndex(0);
  };

  const refreshButtonHandler = () => {
    dispatch(setLoading(true));
    dataset.refresh();
    dispatch(clearChangedRecords());
    dispatch(clearChangedRecordsAfterRefresh());
    dispatch(removeNewRows());
    dispatch(clearInvalidFields());
  };

  const newButtonHandler = () => {
    resetScroll();
    const emptyColumns = columns.map<Column>(column => ({
      schemaName: column.key,
      rawValue: '',
      formattedValue: '',
      type: column.data,
    }));

    dispatch(addNewRow({
      key: Date.now().toString(),
      columns: emptyColumns,
    }));
    _setContainerHeight(getContainerHeight(rows.length + 1));
  };

  const deleteButtonHandler = () => {
    dispatch(setLoading(true));
    dispatch(deleteRecords({ recordIds: selectedRecordIds, _service })).unwrap()
      .then(recordsAfterDelete => {
        dataset.refresh();
        dispatch(readdNewRowsAfterDelete(recordsAfterDelete.newRows));
      })
      .catch(error => {
        if (!error) {
          _service.openErrorDialog(error).then(() => {
            dispatch(setLoading(false));
          });
        }
        dispatch(setLoading(false));
      });
  };

  const saveButtonHandler = () => {
    dispatch(setLoading(true));
    dispatch(saveRecords(_service)).unwrap()
      .then(() => {
        dataset.refresh();
        dispatch(removeNewRows());
      })
      .catch(error =>
        _service.openErrorDialog(error).then(() => {
          dispatch(setLoading(false));
        }));
  };

  React.useEffect(() => {
    const datasetRows = [
      ...newRows,
      ...mapDataSetRows(dataset),
    ];
    dispatch(setRows(datasetRows));
    dispatch(clearChangedRecords());
    dispatch(readdChangedRecordsAfterDelete());
    dispatch(setLoading(isPendingDelete || isPendingLoad));
    _setContainerHeight(getContainerHeight(rows.length));
  }, [dataset, isPendingLoad]);

  useLoadStore(dataset, _service);

  const _renderItemColumn = (item: Row, index: number | undefined, column: IColumn | undefined) =>
    <GridCell row={item} currentColumn={column!} _service={_service} index={index}/>;

  const sort = (sortDirection: ComponentFramework.PropertyHelper.DataSetApi.Types.SortDirection,
    column?: IColumn) => {
    if (column?.fieldName) {
      dispatch(setLoading(true));
      const newSorting: ComponentFramework.PropertyHelper.DataSetApi.SortStatus = {
        name: column.fieldName,
        sortDirection,
      };

      while (dataset.sorting.length > 0) {
        dataset.sorting.pop();
      }
      dataset.sorting.push(newSorting);
      dataset.paging.reset();
      dataset.refresh();
    }
  };

  const onHideSortMenu = React.useCallback(() => setSortMenuProps(undefined), []);

  const getSortMenuProps =
  (ev?: React.MouseEvent<HTMLElement>, column?: IColumn): IContextualMenuProps => {
    const items = [
      { key: 'sortAsc', text: 'Sort Ascending', onClick: () => sort(ASC_SORT, column) },
      { key: 'sortDesc', text: 'Sort Descending', onClick: () => sort(DESC_SORT, column) },
    ];
    return {
      items,
      target: ev?.currentTarget as HTMLElement,
      gapSpace: 2,
      isBeakVisible: false,
      directionalHint: DirectionalHint.bottomLeftEdge,
      onDismiss: onHideSortMenu,
    };
  };

  const _onColumnClick = (ev?: React.MouseEvent<HTMLElement, MouseEvent>, column?: IColumn) => {
    if (column?.columnActionsMode !== ColumnActionsMode.disabled) {
      setSortMenuProps(getSortMenuProps(ev, column));
    }
  };


  // Add effect to filter Invoice lookups
  useEffect(() => {
    const filterInvoiceLookups = async () => {
      console.log('Starting filterInvoiceLookups...');

      const fieldName = 'nb_invoice';
      const parentEntityMetadata = _service.getParentMetadata();
      try {
        console.log('Attempting to get parent record...');
        // Get the supplier code from parent record using webAPI directly
        const parentRecord = await _service.getContext().webAPI.retrieveRecord(
          'nb_ae_chequeregister',
          parentEntityMetadata.entityId,
          '?$select=nb_supplier',
        );
        console.log('Parent Record:', parentRecord);
        const supplierCode = parentRecord?.nb_supplier;
        console.log('Supplier Code:', supplierCode);

        if (supplierCode) {
          console.log('Filtering invoices for supplier:', supplierCode);
          // Filter invoices by supplier code and status
          const filteredInvoices = await _service.retrieveMultipleRecords(
            'nb_ae_invoice',
            `?$select=nb_ae_invoiceid,nb_supplierreference&$filter=
            (nb_supplier eq '${supplierCode}' and nb_invoicestatus eq 124840000)`,
          );
          console.log('Filtered Invoices:', filteredInvoices);

          // Update the lookup options with filtered results
          if (filteredInvoices && filteredInvoices.length > 0) {
            const newFilteredOptions = filteredInvoices.map((invoice: Entity) => {
              console.log('Processing invoice:', invoice);
              const displayName = invoice.nb_supplierreference ||
                `Invoice ${invoice.nb_ae_invoiceid.substring(0, 8)}`;
              console.log('Generated display name:', displayName);
              return {
                key: invoice.nb_ae_invoiceid,
                name: displayName,
                ...invoice,
              };
            });
            console.log('Filtered Options:', newFilteredOptions);
            // Update both the store and local state
            dispatch({
              type: 'lookup/setLookupOptions',
              payload: {
                logicalName: fieldName,
                options: newFilteredOptions,
              },
            });
            console.log('Updated lookup options in store and local state');
          }
          else {
            console.log('No filtered invoices found');
          }
        }
        else {
          console.log('No supplier code found in parent record');
        }
      }
      catch (error: unknown) {
        console.error('Error filtering invoice lookups:', error);
        if (error && typeof error === 'object') {
          const errorObj = error as ErrorDetails;
          console.error('Error details:', {
            code: errorObj.code,
            message: errorObj.message,
            errorCode: errorObj.errorCode,
            title: errorObj.title,
            raw: errorObj.raw,
          });
        }
      }
    };

    filterInvoiceLookups();
  }, []);

  return <div className='container'>
    <Stack horizontal horizontalAlign="end" className={buttonStyles.buttons}>
      <CommandBar
        refreshButtonHandler={refreshButtonHandler}
        newButtonHandler={newButtonHandler}
        deleteButtonHandler={deleteButtonHandler}
        saveButtonHandler={saveButtonHandler}
        isControlDisabled={isControlDisabled}
        selectedCount={selectedRecordIds.length}
      ></CommandBar>
    </Stack>
    <DetailsList
      componentRef={detailsListRef}
      key={getColumnsTotalWidth(dataset) > width ? 0 : width}
      items={rows}
      columns={columns}
      onRenderItemColumn={_renderItemColumn}
      selection={selection}
      onRenderRow={ (props, defaultRender) =>
        <div onDoubleClick={event => {
          const target = event.target as HTMLInputElement;
          if (!target.className.includes('Button')) {
            _service.openForm(props?.item.key);
          }
        }}>
          {defaultRender!(props)}
        </div> }
      onRenderDetailsHeader={_onRenderDetailsHeader}
      layoutMode={DetailsListLayoutMode.fixedColumns}
      styles={gridStyles(rows.length)}
      onColumnHeaderClick={_onColumnClick}
      constrainMode={ ConstrainMode.unconstrained}
    >
    </DetailsList>
    {sortMenuProps && <ContextualMenu {...sortMenuProps} />}
    {rows.length === 0 &&
      <Stack horizontalAlign='center' className='noDataContainer'>
        <div className='nodata'><span>No data available</span></div>
      </Stack>
    }
    <GridFooter
      dataset={dataset}
      selectedCount={selectedRecordIds.length}
      resetScroll={resetScroll}
    ></GridFooter>
  </div>;
};

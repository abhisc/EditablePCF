import React, { useCallback, useState, useEffect } from 'react';
import { IColumn, ITag } from '@fluentui/react';

import { LookupFormat } from '../InputComponents/LookupFormat';
import { NumberFormat } from '../InputComponents/NumberFormat';
import { OptionSetFormat } from '../InputComponents/OptionSetFormat';
import { DateTimeFormat } from '../InputComponents/DateTimeFormat';
import { WholeFormat } from '../InputComponents/WholeFormat';

import { Column, isNewRow, Row } from '../../mappers/dataSetMapper';
import { useAppDispatch, useAppSelector } from '../../store/hooks';
import { updateRow } from '../../store/features/DatasetSlice';
import { setChangedRecords } from '../../store/features/RecordSlice';
import { IDataverseService } from '../../services/DataverseService';
import { TextFormat } from '../InputComponents/TextFormat';

export interface IGridSetProps {
  row: Row,
  currentColumn: IColumn,
  _service: IDataverseService;
  index: number | undefined;
}

export type ParentEntityMetadata = {
  entityId: string,
  entityRecordName: string,
  entityTypeName: string
};

export const GridCell = ({ _service, row, currentColumn, index }: IGridSetProps) => {
  const dispatch = useAppDispatch();
  const cell = row.columns.find((column: Column) => column.schemaName === currentColumn.key);
  const [isInvoiceSelected, setIsInvoiceSelected] = useState(false);

  // Check if this row has an invoice selected
  useEffect(() => {
    const supplierRefCell =
      row.columns.find((column: Column) => column.schemaName === 'nb_supplierreference');
    setIsInvoiceSelected(!!supplierRefCell?.rawValue);
  }, [row.columns]);

  const fieldsRequirementLevels = useAppSelector(state => state.dataset.requirementLevels);
  const fieldRequirementLevel = fieldsRequirementLevels.find(requirementLevel =>
    requirementLevel.fieldName === currentColumn.key);
  const isRequired = fieldRequirementLevel?.isRequired || false;

  const calculatedFields = useAppSelector(state => state.dataset.calculatedFields);
  const calculatedField = calculatedFields.find(field =>
    field.fieldName === currentColumn.key);
  const isCalculatedField = calculatedField?.isCalculated || false;

  const securedFields = useAppSelector(state => state.dataset.securedFields);
  const securedField = securedFields.find(field =>
    field.fieldName === currentColumn.key);
  let hasUpdateAccess = securedField?.hasUpdateAccess || false;

  // Check if this record has been saved
  const savedRecordIds = useAppSelector(state => state.dataset.savedRecordIds);
  const isRecordSaved = savedRecordIds.includes(row.key);

  let parentEntityMetadata: ParentEntityMetadata | undefined;
  let ownerEntityMetadata: string | undefined;
  if (isNewRow(row)) {
    parentEntityMetadata = _service.getParentMetadata();
    ownerEntityMetadata = currentColumn.data === 'Lookup.Owner'
      ? _service.getCurrentUserName() : undefined;
    hasUpdateAccess = securedField?.hasCreateAccess || false;
  }

  const inactiveRecords = useAppSelector(state => state.dataset.inactiveRecords);
  const inactiveRecord = inactiveRecords.find(record =>
    record.recordId === row.key);
  const isInactiveRecord = inactiveRecord?.isInactive || false;

  const _changedValue = useCallback(
    (newValue: unknown, rawValue?: unknown, lookupEntityNavigation?: string): void => {
      dispatch(setChangedRecords({
        id: row.key,
        fieldName: lookupEntityNavigation || currentColumn.key,
        fieldType: currentColumn.data,
        newValue,
      }));

      dispatch(updateRow({
        rowKey: row.key,
        columnName: currentColumn.key,
        newValue: rawValue ?? newValue,
      }));
    }, []);

  const handleInvoiceSelection = async (selected: boolean, invoiceTag?: ITag) => {
    setIsInvoiceSelected(selected);
    if (selected && invoiceTag && invoiceTag.key) {
      try {
        // Fetch invoice details including currency
        const invoice = await _service.getContext().webAPI.retrieveRecord(
          'nb_ae_invoice',
          invoiceTag.key.toString(),
          '?$select=nb_invoice_amt,_transactioncurrencyid_value',
        );
        const dueAmount = invoice.nb_invoice_amt;
        const currencyId = invoice._transactioncurrencyid_value;

        // Update due amount in the row
        dispatch(updateRow({
          rowKey: row.key,
          columnName: 'a_04b6d9baaa2840ac9f6b05c104588d0d.nb_invoice_amt',
          newValue: dueAmount,
        }));

        // Fetch and store currency info for this row
        if (currencyId) {
          const currency = await _service.getCurrencyById(currencyId);
          dispatch({
            type: 'number/addCurrencySymbol',
            payload: {
              recordId: row.key,
              symbol: currency.symbol,
              precision: currency.precision,
            },
          });
        }
      }
      catch (error) {
        console.error('Failed to fetch invoice details or currency:', error);
      }
    }
  };

  const props = {
    fieldName: currentColumn?.fieldName ? currentColumn?.fieldName : '',
    fieldId: `${currentColumn?.fieldName || ''}${row.key}`,
    formattedValue: cell?.formattedValue,
    isRequired,
    isDisabled: isInactiveRecord || isCalculatedField || isRecordSaved ||
      (!isInvoiceSelected && currentColumn.key !== 'nb_supplierreference' &&
        currentColumn.key !== 'nb_invoice_posting_amt'),
    isSecured: !hasUpdateAccess,
    _onChange: _changedValue,
    _service,
    index,
    ownerValue: ownerEntityMetadata,
  };

  if (currentColumn !== undefined && cell !== undefined) {
    switch (currentColumn.data) {
      case 'DateAndTime.DateAndTime':
        return <DateTimeFormat dateOnly={false} value={cell.rawValue} rowId={row.key} {...props} />;

      case 'DateAndTime.DateOnly':
        return <DateTimeFormat dateOnly={true} value={cell.rawValue} rowId={row.key} {...props} />;

      case 'Lookup.Simple':
        return <LookupFormat
          value={cell.lookup}
          parentEntityMetadata={parentEntityMetadata}
          onInvoiceSelected={
            currentColumn.key === 'nb_supplierreference' ? handleInvoiceSelection : undefined
          }
          rowId={row.key}
          {...props}
        />;

      case 'Lookup.Customer':
      case 'Lookup.Owner':
        return <TextFormat
          value={cell.formattedValue}
          rowId={row.key}
          {...props}
          isDisabled={true}
        />;

      case 'OptionSet':
        return <OptionSetFormat
          value={cell.rawValue}
          isMultiple={false}
          rowId={row.key}
          {...props}
        />;

      case 'TwoOptions':
        return <OptionSetFormat
          value={cell.rawValue}
          isMultiple={false}
          isTwoOptions={true}
          rowId={row.key}
          {...props}
        />;

      case 'MultiSelectPicklist':
        return <OptionSetFormat
          value={cell.rawValue}
          isMultiple={true}
          rowId={row.key}
          {...props}
        />;

      case 'Decimal':
        return <NumberFormat value={cell.formattedValue ?? ''} {...props} />;

      case 'Currency':
        return <NumberFormat value={cell.formattedValue ?? ''} {...props} />;

      case 'FP':
        return <NumberFormat value={cell.formattedValue ?? ''} {...props} />;

      case 'Whole.None':
        return <NumberFormat value={cell.formattedValue ?? ''} {...props} />;

      case 'Whole.Duration':
        return <WholeFormat
          value={cell.rawValue}
          type={'duration'}
          rowId={row.key}
          {...props}
        />;

      case 'Whole.Language':
        return <WholeFormat
          value={cell.rawValue}
          type={'language'}
          rowId={row.key}
          {...props}
        />;

      case 'Whole.TimeZone':
        return <WholeFormat
          value={cell.rawValue}
          type={'timezone'}
          rowId={row.key}
          {...props}
        />;

      case 'SingleLine.Text':
      case 'Multiple':
      default:
        return <TextFormat
          value={cell.formattedValue || ''}
          type={cell.type}
          rowId={row.key}
          {...props}
        />;
    }
  }

  return <></>;
};

/* eslint-disable react/display-name */
import { DefaultButton, FontIcon } from '@fluentui/react';
import { ITag, TagPicker } from '@fluentui/react/lib/Pickers';
import React, { memo } from 'react';
import { IDataverseService } from '../../services/DataverseService';
import { useAppDispatch, useAppSelector } from '../../store/hooks';
import {
  asteriskClassStyle,
  lookupFormatStyles,
  lookupSelectedOptionStyles,
} from '../../styles/ComponentsStyles';
import { ParentEntityMetadata } from '../EditableGrid/GridCell';
import { ErrorIcon } from '../ErrorIcon';
import { setInvalidFields } from '../../store/features/ErrorSlice';

const MAX_NUMBER_OF_OPTIONS = 100;
const SINGLE_CLICK_CODE = 1;

export interface ILookupProps {
  fieldId: string;
  fieldName: string;
  value: ITag | undefined;
  parentEntityMetadata: ParentEntityMetadata | undefined;
  isRequired: boolean;
  isSecured: boolean;
  isDisabled: boolean;
  _onChange: Function;
  _service: IDataverseService;
  onInvoiceSelected?: (isSelected: boolean) => void;
}

export const LookupFormat = memo(({ fieldId, fieldName, value, parentEntityMetadata,
  isSecured, isRequired, isDisabled, _onChange, _service, onInvoiceSelected }: ILookupProps) => {
  const picker = React.useRef(null);
  const dispatch = useAppDispatch();

  const lookups = useAppSelector(state => state.lookup.lookups);
  const currentLookup = lookups.find(lookup => lookup.logicalName === fieldName);
  const options = currentLookup?.options ?? [];
  const currentOption = value ? [value] : [];
  const isOffline = _service.isOffline();

  // Add state to track filtered options
  const filteredOptions = [...options];

  if (value === undefined &&
    parentEntityMetadata !== undefined && parentEntityMetadata.entityId !== undefined) {
    if (currentLookup?.reference?.entityNameRef === parentEntityMetadata.entityTypeName) {
      currentOption.push({
        key: parentEntityMetadata.entityId,
        name: parentEntityMetadata.entityRecordName,
      });

      _onChange(`/${currentLookup?.entityPluralName}(${parentEntityMetadata.entityId})`,
        currentOption[0],
        currentLookup?.reference?.entityNavigation);
    }
  }

  const initialValues = (): ITag[] => {
    const optionsToUse = fieldName === 'nb_invoice' ? filteredOptions : options;
    if (optionsToUse.length > MAX_NUMBER_OF_OPTIONS) {
      return optionsToUse.slice(0, MAX_NUMBER_OF_OPTIONS);
    }
    return optionsToUse;
  };

  const filterSuggestedTags = (filterText: string): ITag[] => {
    if (filterText.length === 0) return [];

    const optionsToUse = fieldName === 'nb_invoice' ? filteredOptions : options;
    return optionsToUse.filter(tag => {
      if (tag.name === null) return false;
      return tag.name.toLowerCase().includes(filterText.toLowerCase());
    });
  };

  const onChange = (items?: ITag[] | undefined): void => {
    if (items !== undefined && items.length > 0) {
      _onChange(`/${currentLookup?.entityPluralName}(${items[0].key})`, items[0],
        currentLookup?.reference?.entityNavigation);
      if (fieldName === 'nb_invoice' && onInvoiceSelected) {
        onInvoiceSelected(true);
      }
    }
    else {
      _onChange(null, null, currentLookup?.reference?.entityNavigation);
      if (fieldName === 'nb_invoice' && onInvoiceSelected) {
        onInvoiceSelected(false);
      }
    }
  };

  const _onRenderItem = () =>
    <DefaultButton
      text={currentOption[0].name}
      title={currentOption[0].name}
      menuProps={{ items: [] }}
      split
      menuIconProps={{
        iconName: 'Cancel',
      }}
      onMenuClick={() => onChange(undefined)}
      onClick={event => {
        if (event.detail === SINGLE_CLICK_CODE) {
          _service.openForm(currentOption[0].key.toString(),
            currentLookup?.reference?.entityNameRef);
        }
      }}
      styles={lookupSelectedOptionStyles}
    />;

  // Determine if this field should be editable
  const isEditable = fieldName === 'nb_invoice' || value !== undefined;

  return <div>
    <TagPicker
      selectedItems={currentOption}
      componentRef={picker}
      onChange={onChange}
      onResolveSuggestions={filterSuggestedTags}
      resolveDelay={1000}
      onEmptyResolveSuggestions={initialValues}
      itemLimit={1}
      pickerSuggestionsProps={{ noResultsFoundText: 'No Results Found' }}
      styles={lookupFormatStyles(isRequired, isSecured ||
        (!isEditable && fieldName !== 'nb_invoice') || isDisabled || isOffline)}
      onRenderItem={ !isSecured && isEditable &&
        !isDisabled && !isOffline ? _onRenderItem : undefined}
      onBlur={() => {
        if (picker.current) {
          // eslint-disable-next-line @typescript-eslint/ban-ts-comment
          // @ts-ignore
          picker.current.input.current._updateValue('');
          dispatch(setInvalidFields({ fieldId, isInvalid: isRequired,
            errorMessage: 'Required fields must be filled in' }));
        }
      }}
      disabled={isSecured || (!isEditable && fieldName !== 'nb_invoice') || isDisabled || isOffline}
      inputProps={{
        onFocus: () => dispatch(setInvalidFields({ fieldId, isInvalid: false, errorMessage: '' })),
      }}
    />
    <FontIcon iconName={'AsteriskSolid'} className={asteriskClassStyle(isRequired)} />
    <ErrorIcon id={fieldId} isRequired={isRequired} />
  </div>;
});

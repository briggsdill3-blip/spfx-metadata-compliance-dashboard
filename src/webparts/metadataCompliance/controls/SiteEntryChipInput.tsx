import * as React from 'react';
import { useState } from 'react';
import styles from './SiteEntryChipInput.module.scss';
import { ISiteEntry, deriveLabel } from '../components/ISiteEntry';

export interface ISiteEntryChipInputProps {
  label: string;
  entries: ISiteEntry[];
  onChange: (entries: ISiteEntry[]) => void;
}

const SiteEntryChipInput: React.FunctionComponent<ISiteEntryChipInputProps> = (props) => {
  const [inputValue, setInputValue] = useState<string>('');

  const addEntry = (): void => {
    const trimmed = inputValue.trim();
    if (trimmed.length === 0) {
      return;
    }

    const alreadyExists = props.entries.some(
      (entry) => entry.url.toLowerCase() === trimmed.toLowerCase()
    );
    if (alreadyExists) {
      setInputValue('');
      return;
    }

    const newEntry: ISiteEntry = {
      url: trimmed,
      label: deriveLabel(trimmed)
    };

    props.onChange([...props.entries, newEntry]);
    setInputValue('');
  };

  const removeEntry = (url: string): void => {
    props.onChange(props.entries.filter((entry) => entry.url !== url));
  };

  const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>): void => {
    if (e.key === 'Enter') {
      e.preventDefault();
      addEntry();
    }
  };

  return (
    <div className={styles.chipInputWrapper}>
      <div className={styles.fieldLabel}>{props.label}</div>

      <div className={styles.inputRow}>
        <input
          type="text"
          className={styles.textInput}
          placeholder="Paste a site URL and press Enter"
          value={inputValue}
          onChange={(e) => setInputValue(e.target.value)}
          onKeyDown={handleKeyDown}
        />
        <button
          type="button"
          className={styles.addButton}
          onClick={addEntry}
          disabled={inputValue.trim().length === 0}
        >
          Add
        </button>
      </div>

      {props.entries.length > 0 && (
        <ul className={styles.chipList}>
          {props.entries.map((entry) => (
            <li key={entry.url} className={styles.chip} title={entry.url}>
              <span className={styles.chipLabel}>{entry.label}</span>
              <button
                type="button"
                className={styles.chipRemove}
                onClick={() => removeEntry(entry.url)}
                aria-label={`Remove ${entry.label}`}
              >
                ×
              </button>
            </li>
          ))}
        </ul>
      )}
    </div>
  );
};

export default SiteEntryChipInput;
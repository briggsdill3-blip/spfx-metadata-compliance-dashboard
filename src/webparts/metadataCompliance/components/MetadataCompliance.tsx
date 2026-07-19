import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './MetadataCompliance.module.scss';
import type { IMetadataComplianceProps } from './IMetadataComplianceProps';

interface ILibraryOption {
  Title: string;
  ItemCount: number;
}

interface IFieldMeta {
  InternalName: string;
  Title: string;
  TypeAsString: string;
}

interface IRawFieldResponse {
  InternalName: string;
  Title: string;
  TypeAsString: string;
  Hidden: boolean;
  ReadOnlyField: boolean;
  Group: string;
}

interface IDocItem {
  Id: number;
  FileLeafRef: string;
  values: Record<string, string>;
}

const SYSTEM_FIELD_BLOCKLIST = new Set([
  'ContentType', 'Created', 'Author', 'Editor', 'Modified', 'UIVersionString',
  'Attachments', 'Edit', 'LinkTitleNoMenu', 'LinkTitle', 'DocIcon',
  'ItemChildCount', 'FolderChildCount', 'AppAuthor', 'AppEditor',
  'owshiddenversion', 'WorkflowVersion', 'WorkflowInstanceID',
  'FileRef', 'FileDirRef', 'FSObjType', 'SortBehavior', 'PermMask', 'UniqueId',
  'SyncClientId', 'ProgId', 'ScopeId', 'MetaInfo', 'InstanceID', 'Order', 'GUID',
  'CheckedOutTitle', 'CheckedOutUserId', 'IsCheckedoutToLocal', 'VirusStatus',
  'TemplateUrl', 'ParentVersionString', 'ParentLeafName',
  'FileLeafRef', 'Title'
]);

const isFieldEmpty = (value: string): boolean => {
  return !value || value.trim() === '';
};

const isItemComplete = (item: IDocItem, fieldTitles: string[]): boolean => {
  return fieldTitles.every(title => !isFieldEmpty(item.values[title]));
};

const getStatusTier = (percent: number, goodThreshold: number, warnThreshold: number): 'good' | 'warn' | 'bad' => {
  if (percent >= goodThreshold) return 'good';
  if (percent >= warnThreshold) return 'warn';
  return 'bad';
};

interface IProgressRingProps {
  percent: number;
  tier: 'good' | 'warn' | 'bad';
}

const ProgressRing: React.FunctionComponent<IProgressRingProps> = ({ percent, tier }) => {
  const [displayPercent, setDisplayPercent] = useState<number>(0);

  useEffect(() => {
    setDisplayPercent(0);
    const frame1 = requestAnimationFrame(() => {
      const frame2 = requestAnimationFrame(() => setDisplayPercent(percent));
      return () => cancelAnimationFrame(frame2);
    });
    return () => cancelAnimationFrame(frame1);
  }, [percent]);

  const radius = 70;
  const stroke = 12;
  const normalizedRadius = radius - stroke / 2;
  const circumference = normalizedRadius * 2 * Math.PI;
  const offset = circumference - (displayPercent / 100) * circumference;

  return (
    <svg height={radius * 2} width={radius * 2} className={styles.ring}>
      <circle
        stroke="currentColor"
        className={styles.ringTrack}
        fill="transparent"
        strokeWidth={stroke}
        r={normalizedRadius}
        cx={radius}
        cy={radius}
      />
      <circle
        stroke="currentColor"
        className={`${styles.ringProgress} ${styles[`tier-${tier}`]}`}
        fill="transparent"
        strokeWidth={stroke}
        strokeDasharray={`${circumference} ${circumference}`}
        style={{ strokeDashoffset: offset }}
        strokeLinecap="round"
        r={normalizedRadius}
        cx={radius}
        cy={radius}
      />
      <text x="50%" y="50%" textAnchor="middle" dy="0.35em" className={styles.ringLabel}>
        {percent}%
      </text>
    </svg>
  );
};

const getCustomFields = async (
  sp: IMetadataComplianceProps['sp'],
  libraryTitle: string,
  excludedFields: string[]
): Promise<IFieldMeta[]> => {
  const rawFields = await sp.web.lists.getByTitle(libraryTitle).fields
    .select('InternalName', 'Title', 'TypeAsString', 'Hidden', 'ReadOnlyField', 'Group')
    .filter('Hidden eq false and ReadOnlyField eq false')();

  return rawFields
    .filter((f: IRawFieldResponse) =>
      !SYSTEM_FIELD_BLOCKLIST.has(f.InternalName) &&
      !f.InternalName.startsWith('_') &&
      f.Group !== '_Hidden' &&
      f.Group !== 'Base Columns' &&
      excludedFields.indexOf(f.Title.toLowerCase()) === -1
    )
    .map((f: IRawFieldResponse) => ({
      InternalName: f.InternalName,
      Title: f.Title,
      TypeAsString: f.TypeAsString
    }));
};

const MetadataCompliance: React.FunctionComponent<IMetadataComplianceProps> = (props) => {
  const [libraries, setLibraries] = useState<ILibraryOption[]>([]);
  const [selectedLibrary, setSelectedLibrary] = useState<string>('');
  const [librariesLoading, setLibrariesLoading] = useState<boolean>(true);
  const [fieldsCache, setFieldsCache] = useState<Record<string, IFieldMeta[]>>({});

  const [items, setItems] = useState<IDocItem[]>([]);
  const [dataLoading, setDataLoading] = useState<boolean>(false);
  const [error, setError] = useState<string>('');
  const [selectedType, setSelectedType] = useState<string>('All');

  const isLocked = props.lockedLibrary.trim().length > 0;
  const excludedKey = props.excludedFields.join('|');

  useEffect(() => {
    const setupLockedLibrary = async (): Promise<void> => {
      try {
        const libTitle = props.lockedLibrary.trim();
        const rawLib = await props.sp.web.lists.getByTitle(libTitle).select('Title', 'ItemCount')();
        const customFields = await getCustomFields(props.sp, libTitle, props.excludedFields);

        setFieldsCache({ [libTitle]: customFields });
        setSelectedLibrary(libTitle);
        setLibraries([{ Title: libTitle, ItemCount: rawLib.ItemCount }]);
      } catch (err) {
        setError(`Unable to find or read the configured library "${props.lockedLibrary}".`);
        console.error(err);
      } finally {
        setLibrariesLoading(false);
      }
    };

    const discoverLibraries = async (): Promise<void> => {
      try {
        const rawLibraries = await props.sp.web.lists
          .filter('BaseTemplate eq 101 and Hidden eq false')
          .select('Title', 'ItemCount')();

        const candidateLibraries: ILibraryOption[] = rawLibraries.map((lib: ILibraryOption) => ({
          Title: lib.Title,
          ItemCount: lib.ItemCount
        }));

        const fieldResults = await Promise.all(
          candidateLibraries.map(async lib => ({
            title: lib.Title,
            fields: await getCustomFields(props.sp, lib.Title, props.excludedFields)
          }))
        );

        const cache: Record<string, IFieldMeta[]> = {};
        fieldResults.forEach(r => { cache[r.title] = r.fields; });
        setFieldsCache(cache);

        const qualifyingLibraries = candidateLibraries
          .filter(lib => cache[lib.Title] && cache[lib.Title].length > 0)
          .sort((a, b) => a.Title.localeCompare(b.Title));

        setLibraries(qualifyingLibraries);

        const defaultLib = qualifyingLibraries.some(l => l.Title === 'Documents')
          ? 'Documents'
          : (qualifyingLibraries[0] ? qualifyingLibraries[0].Title : '');

        setSelectedLibrary(defaultLib);
      } catch (err) {
        setError('Unable to load document libraries for this site.');
        console.error(err);
      } finally {
        setLibrariesLoading(false);
      }
    };

    setLibrariesLoading(true);
    if (isLocked) {
      setupLockedLibrary().catch((err) => console.error(err));
    } else {
      discoverLibraries().catch((err) => console.error(err));
    }
  }, [props.sp, props.lockedLibrary, excludedKey]);

  useEffect(() => {
    if (!selectedLibrary || !fieldsCache[selectedLibrary]) {
      return;
    }

    const loadItems = async (): Promise<void> => {
      setDataLoading(true);
      setError('');
      setSelectedType('All');

      const libraryFields = fieldsCache[selectedLibrary];

      try {
        const selectList = ['Id', 'FileLeafRef', ...libraryFields.map(f =>
          f.TypeAsString === 'User' ? `${f.InternalName}/Title` : f.InternalName
        )];
        const expandList = libraryFields.filter(f => f.TypeAsString === 'User').map(f => f.InternalName);

        let query = props.sp.web.lists.getByTitle(selectedLibrary).items
          .select(...selectList)
          .top(5000);

        if (expandList.length > 0) {
          query = query.expand(...expandList);
        }

        const rawItems = await query();

        const mapped: IDocItem[] = rawItems.map((raw: Record<string, unknown>) => {
          const values: Record<string, string> = {};
          libraryFields.forEach(f => {
            const rawUserValue = raw[f.InternalName] as { Title?: string } | undefined;
            const rawValue = f.TypeAsString === 'User'
              ? (rawUserValue && rawUserValue.Title ? rawUserValue.Title : '')
              : (raw[f.InternalName] as string | string[] | undefined);

            if (Array.isArray(rawValue)) {
              values[f.Title] = rawValue.join(', ');
            } else {
              values[f.Title] = rawValue || '';
            }
          });

          return {
            Id: raw.Id as number,
            FileLeafRef: raw.FileLeafRef as string,
            values
          };
        });

        setItems(mapped);
      } catch (err) {
        setError(`Unable to load data from "${selectedLibrary}".`);
        console.error(err);
      } finally {
        setDataLoading(false);
      }
    };

    loadItems().catch((err) => console.error(err));
  }, [props.sp, selectedLibrary, fieldsCache]);

  const currentFields = fieldsCache[selectedLibrary] || [];
  const keyFieldTitles = currentFields.map(f => f.Title);

  const typeField = currentFields.find(f => f.Title.toLowerCase().includes('type'));
  const documentTypes = typeField
    ? ['All', ...Array.from(new Set(items.map(i => i.values[typeField.Title]).filter(t => t !== '')))]
    : ['All'];

  const filteredItems = (!typeField || selectedType === 'All')
    ? items
    : items.filter(i => i.values[typeField.Title] === selectedType);

  const totalCount = filteredItems.length;
  const completeCount = filteredItems.filter(i => isItemComplete(i, keyFieldTitles)).length;
  const incompleteCount = totalCount - completeCount;
  const completionPercent = totalCount === 0 ? 0 : Math.round((completeCount / totalCount) * 100);
  const tier = getStatusTier(completionPercent, props.goodThreshold, props.warnThreshold);

  const missingByField = keyFieldTitles.reduce((acc, title) => {
    acc[title] = filteredItems.filter(i => isFieldEmpty(i.values[title])).length;
    return acc;
  }, {} as Record<string, number>);

  const maxMissing = Math.max(1, ...keyFieldTitles.map(t => missingByField[t]));
  const isBusy = librariesLoading || dataLoading;

  if (librariesLoading) {
    return (
      <section className={`${styles.metadataCompliance} ${props.isDarkTheme ? styles.dark : ''}`}>
        <div className={styles.loadingState}>
          <div className={styles.spinner} />
          <span>Discovering document libraries...</span>
        </div>
      </section>
    );
  }

  if (libraries.length === 0 && !error) {
    return (
      <section className={`${styles.metadataCompliance} ${props.isDarkTheme ? styles.dark : ''}`}>
        <div className={styles.errorState}>
          No document libraries with custom metadata columns were found on this site.
        </div>
      </section>
    );
  }

  return (
    <section className={`${styles.metadataCompliance} ${props.isDarkTheme ? styles.dark : ''}`}>
      <header className={styles.header}>
        <div>
          <h2 className={styles.title}>Metadata Compliance Dashboard</h2>
          <p className={styles.subtitle}>Tagging health for this document library</p>
        </div>
        <div className={styles.filterGroup}>
          {!isLocked && (
            <div className={styles.filterControl}>
              <label htmlFor="libraryFilter" className={styles.filterLabel}>Library</label>
              <select
                id="libraryFilter"
                className={styles.select}
                value={selectedLibrary}
                onChange={(e) => setSelectedLibrary(e.target.value)}
              >
                {libraries.map(lib => (
                  <option key={lib.Title} value={lib.Title}>
                    {lib.Title} ({lib.ItemCount})
                  </option>
                ))}
              </select>
            </div>
          )}

          {typeField && (
            <div className={styles.filterControl}>
              <label htmlFor="typeFilter" className={styles.filterLabel}>{typeField.Title}</label>
              <select
                id="typeFilter"
                className={styles.select}
                value={selectedType}
                onChange={(e) => setSelectedType(e.target.value)}
                disabled={isBusy || items.length === 0}
              >
                {documentTypes.map(type => (
                  <option key={type} value={type}>{type}</option>
                ))}
              </select>
            </div>
          )}
        </div>
      </header>

      {isBusy && (
        <div className={styles.loadingState}>
          <div className={styles.spinner} />
          <span>Loading library data...</span>
        </div>
      )}

      {!isBusy && error && (
        <div className={styles.errorState}>{error}</div>
      )}

      {!isBusy && !error && (
        <>
          <div className={styles.summaryRow}>
            <div className={styles.ringCard}>
              <ProgressRing percent={completionPercent} tier={tier} />
              <span className={`${styles.tierBadge} ${styles[`tier-${tier}`]}`}>
                {tier === 'good' ? 'On Target' : tier === 'warn' ? 'Needs Attention' : 'At Risk'}
              </span>
            </div>

            <div className={styles.statCards}>
              <div className={styles.statCard}>
                <span className={styles.statValue}>{totalCount}</span>
                <span className={styles.statLabel}>Total Items</span>
              </div>
              <div className={`${styles.statCard} ${styles.statGood}`}>
                <span className={styles.statValue}>{completeCount}</span>
                <span className={styles.statLabel}>Fully Tagged</span>
              </div>
              <div className={`${styles.statCard} ${styles.statBad}`}>
                <span className={styles.statValue}>{incompleteCount}</span>
                <span className={styles.statLabel}>Incomplete Items</span>
              </div>
            </div>
          </div>

          <div className={styles.breakdown}>
            <h3 className={styles.breakdownTitle}>Missing Fields Breakdown</h3>
            <div className={styles.barList}>
              {keyFieldTitles.map(title => {
                const missing = missingByField[title];
                const widthPercent = (missing / maxMissing) * 100;
                return (
                  <div className={styles.barRow} key={title}>
                    <span className={styles.barLabel}>{title}</span>
                    <div className={styles.barTrack}>
                      <div
                        className={styles.barFill}
                        style={{ width: `${widthPercent}%` }}
                      />
                    </div>
                    <span className={styles.barCount}>{missing}</span>
                  </div>
                );
              })}
            </div>
          </div>
        </>
      )}
    </section>
  );
};

export default MetadataCompliance;
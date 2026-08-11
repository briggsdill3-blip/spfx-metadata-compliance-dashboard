import * as React from 'react';
import { useState, useEffect } from 'react';
import { Web } from '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import styles from './MetadataCompliance.module.scss';
import type { IMetadataComplianceProps } from './IMetadataComplianceProps';
import type { SPFI } from '@pnp/sp';

interface ILibraryOption {
  siteUrl: string;
  siteLabel: string;
  title: string;
  itemCount: number;
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

const ALL_SITES_KEY = '__all__';

const makeLibraryKey = (siteUrl: string, title: string): string => `${siteUrl}::${title}`;

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

const getTierClassName = (tier: 'good' | 'warn' | 'bad'): string => {
  const tierKey = `tier-${tier}`;
  return styles[tierKey];
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

  const radius = 60;
  const stroke = 10;
  const normalizedRadius = radius - stroke / 2;
  const circumference = normalizedRadius * 2 * Math.PI;
  const offset = circumference - (displayPercent / 100) * circumference;
  const tierClass = getTierClassName(tier);

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
        className={`${styles.ringProgress} ${tierClass}`}
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
  web: ReturnType<typeof Web>,
  libraryTitle: string,
  excludedFields: string[]
): Promise<IFieldMeta[]> => {
  const rawFields = await web.lists.getByTitle(libraryTitle).fields
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
  const [selectedKey, setSelectedKey] = useState<string>('');
  const [selectedSiteFilter, setSelectedSiteFilter] = useState<string>(ALL_SITES_KEY);
  const [librariesLoading, setLibrariesLoading] = useState<boolean>(true);
  const [fieldsCache, setFieldsCache] = useState<Record<string, IFieldMeta[]>>({});

  const [items, setItems] = useState<IDocItem[]>([]);
  const [dataLoading, setDataLoading] = useState<boolean>(false);
  const [error, setError] = useState<string>('');
  const [selectedType, setSelectedType] = useState<string>('All');

  const sitesKey = props.targetSites.map(s => s.url).join('|');
  const excludedLibrariesKey = props.excludedLibraries.join('|');
  const excludedFieldsKey = props.excludedFields.join('|');

  useEffect(() => {
    let cancelled = false;

    const discover = async (): Promise<void> => {
      setLibrariesLoading(true);
      setError('');

      try {
        const perSite = await Promise.all(
          props.targetSites.map(async (site) => {
            try {
              const web = Web([props.sp.web, site.url]);
              const rawLibraries = await web.lists
                .filter('BaseTemplate eq 101 and Hidden eq false')
                .select('Title', 'ItemCount')();
              return { site, libraries: rawLibraries as { Title: string; ItemCount: number }[] };
            } catch (err) {
              console.error(`Failed to load libraries for ${site.url}`, err);
              return { site, libraries: [] as { Title: string; ItemCount: number }[] };
            }
          })
        );

        const candidates: ILibraryOption[] = [];
        perSite.forEach(({ site, libraries: siteLibraries }) => {
          siteLibraries.forEach((lib) => {
            const key = makeLibraryKey(site.url, lib.Title);
            if (props.excludedLibraries.indexOf(key) === -1) {
              candidates.push({
                siteUrl: site.url,
                siteLabel: site.label,
                title: lib.Title,
                itemCount: lib.ItemCount
              });
            }
          });
        });

        const fieldResults = await Promise.all(
          candidates.map(async (c) => {
            const web = Web([props.sp.web, c.siteUrl]);
            const fields = await getCustomFields(web, c.title, props.excludedFields);
            return { key: makeLibraryKey(c.siteUrl, c.title), fields };
          })
        );

        const cache: Record<string, IFieldMeta[]> = {};
        fieldResults.forEach((r) => { cache[r.key] = r.fields; });

        const qualifying = candidates
          .filter((c) => cache[makeLibraryKey(c.siteUrl, c.title)] && cache[makeLibraryKey(c.siteUrl, c.title)].length > 0)
          .sort((a, b) => a.siteLabel.localeCompare(b.siteLabel) || a.title.localeCompare(b.title));

        if (!cancelled) {
          setFieldsCache(cache);
          setLibraries(qualifying);
          setSelectedKey(qualifying.length > 0 ? makeLibraryKey(qualifying[0].siteUrl, qualifying[0].title) : '');
        }
      } catch (err) {
        if (!cancelled) {
          setError('Unable to load document libraries across the configured sites.');
        }
        console.error(err);
      } finally {
        if (!cancelled) {
          setLibrariesLoading(false);
        }
      }
    };

    if (props.targetSites.length === 0) {
      setLibraries([]);
      setSelectedKey('');
      setLibrariesLoading(false);
    } else {
      discover().catch((err) => console.error(err));
    }

    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [props.sp, sitesKey, excludedLibrariesKey, excludedFieldsKey]);

  useEffect(() => {
    if (!selectedKey || !fieldsCache[selectedKey]) {
      return;
    }

    const selectedLib = libraries.find((l) => makeLibraryKey(l.siteUrl, l.title) === selectedKey);
    if (!selectedLib) {
      return;
    }

    const loadItems = async (): Promise<void> => {
      setDataLoading(true);
      setError('');
      setSelectedType('All');

      const libraryFields = fieldsCache[selectedKey];
      const web = Web([props.sp.web, selectedLib.siteUrl]);

      try {
        const selectList = ['Id', 'FileLeafRef', ...libraryFields.map(f =>
          f.TypeAsString === 'User' ? `${f.InternalName}/Title` : f.InternalName
        )];
        const expandList = libraryFields.filter(f => f.TypeAsString === 'User').map(f => f.InternalName);

        let query = web.lists.getByTitle(selectedLib.title).items
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
        setError(`Unable to load data from "${selectedLib.title}".`);
        console.error(err);
      } finally {
        setDataLoading(false);
      }
    };

    loadItems().catch((err) => console.error(err));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [props.sp, selectedKey, fieldsCache]);

  const currentFields = fieldsCache[selectedKey] || [];
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
  const tierClass = getTierClassName(tier);

  const missingByField = keyFieldTitles.reduce((acc, title) => {
    acc[title] = filteredItems.filter(i => isFieldEmpty(i.values[title])).length;
    return acc;
  }, {} as Record<string, number>);

  const maxMissing = Math.max(1, ...keyFieldTitles.map(t => missingByField[t]));
  const isBusy = librariesLoading || dataLoading;

  const uniqueSites = Array.from(new Map(libraries.map(l => [l.siteUrl, l.siteLabel])).entries());

  const librariesForSiteFilter = selectedSiteFilter === ALL_SITES_KEY
    ? libraries
    : libraries.filter(l => l.siteUrl === selectedSiteFilter);

  const themeColors = props.theme ? props.theme.semanticColors : undefined;
  const themePalette = props.theme ? props.theme.palette : undefined;

  const rootStyle: React.CSSProperties = themeColors && themePalette ? ({
    '--mcd-text': themeColors.bodyText,
    '--mcd-text-secondary': themeColors.bodySubtext || themeColors.bodyText,
    '--mcd-bg-surface': themeColors.bodyBackground,
    '--mcd-bg-hover': themeColors.bodyBackgroundHovered || themeColors.bodyBackground,
    '--mcd-border': themeColors.bodyDivider || themePalette.neutralLight,
    '--mcd-accent': themePalette.themePrimary,
    '--mcd-accent-text': themePalette.white
  } as React.CSSProperties) : {};

  if (props.targetSites.length === 0) {
    return (
      <section className={styles.metadataCompliance} style={rootStyle}>
        <div className={styles.errorState}>
          This web part needs to be configured. Open the edit panel and add one or more target sites.
        </div>
      </section>
    );
  }

  if (librariesLoading) {
    return (
      <section className={styles.metadataCompliance} style={rootStyle}>
        <div className={styles.loadingState}>
          <div className={styles.spinner} />
          <span>Discovering document libraries...</span>
        </div>
      </section>
    );
  }

  if (libraries.length === 0 && !error) {
    return (
      <section className={styles.metadataCompliance} style={rootStyle}>
        <div className={styles.errorState}>
          No qualifying document libraries were found across the configured sites.
        </div>
      </section>
    );
  }

  return (
    <section className={styles.metadataCompliance} style={rootStyle}>
      <header className={styles.header}>
        <div>
          <h2 className={styles.title}>Metadata Compliance Dashboard</h2>
          <p className={styles.subtitle}>Tagging health across your configured sites</p>
        </div>
        <div className={styles.filterGroup}>
          <div className={styles.filterControl}>
            <label htmlFor="siteFilter" className={styles.filterLabel}>Site</label>
            <select
              id="siteFilter"
              className={styles.select}
              value={selectedSiteFilter}
              onChange={(e) => setSelectedSiteFilter(e.target.value)}
            >
              <option value={ALL_SITES_KEY}>All Sites</option>
              {uniqueSites.map(([url, label]) => (
                <option key={url} value={url}>{label}</option>
              ))}
            </select>
          </div>

          <div className={styles.filterControl}>
            <label htmlFor="libraryFilter" className={styles.filterLabel}>Library</label>
            <select
              id="libraryFilter"
              className={styles.select}
              value={selectedKey}
              onChange={(e) => setSelectedKey(e.target.value)}
            >
              {librariesForSiteFilter.map(lib => {
                const key = makeLibraryKey(lib.siteUrl, lib.title);
                const displayLabel = selectedSiteFilter === ALL_SITES_KEY
                  ? `${lib.siteLabel}: ${lib.title} (${lib.itemCount})`
                  : `${lib.title} (${lib.itemCount})`;
                return (
                  <option key={key} value={key}>{displayLabel}</option>
                );
              })}
            </select>
          </div>

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
              <span className={`${styles.tierBadge} ${tierClass}`}>
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
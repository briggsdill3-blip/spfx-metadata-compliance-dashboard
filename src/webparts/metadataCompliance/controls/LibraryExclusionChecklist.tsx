import * as React from 'react';
import { useState, useEffect } from 'react';
import { SPFI } from '@pnp/sp';
import { Web } from '@pnp/sp/webs';
import '@pnp/sp/lists';
import styles from './LibraryExclusionChecklist.module.scss';
import type { ISiteEntry } from '../components/ISiteEntry';

interface ILibraryRow {
  siteUrl: string;
  title: string;
  itemCount: number;
}

interface ISiteGroup {
  siteUrl: string;
  siteLabel: string;
  libraries: ILibraryRow[];
  error: string;
}

export interface ILibraryExclusionChecklistProps {
  label: string;
  sp: SPFI;
  sites: ISiteEntry[];
  excludedLibraries: string[];
  onChange: (excludedLibraries: string[]) => void;
}

const makeKey = (siteUrl: string, title: string): string => `${siteUrl}::${title}`;

const LibraryExclusionChecklist: React.FunctionComponent<ILibraryExclusionChecklistProps> = (props) => {
  const [groups, setGroups] = useState<ISiteGroup[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const siteUrlsKey = props.sites.map((s) => s.url).join('|');

  useEffect(() => {
    let cancelled = false;

    const loadLibraries = async (): Promise<void> => {
      setLoading(true);

      const results: ISiteGroup[] = await Promise.all(
        props.sites.map(async (site): Promise<ISiteGroup> => {
          try {
            const web = Web([props.sp.web, site.url]);
            const rawLibraries = await web.lists
              .filter('BaseTemplate eq 101 and Hidden eq false')
              .select('Title', 'ItemCount')();

            const libraries: ILibraryRow[] = (rawLibraries as { Title: string; ItemCount: number }[])
              .map((lib) => ({
                siteUrl: site.url,
                title: lib.Title,
                itemCount: lib.ItemCount
              }))
              .sort((a, b) => a.title.localeCompare(b.title));

            return { siteUrl: site.url, siteLabel: site.label, libraries, error: '' };
          } catch (err) {
            console.error(`Failed to load libraries for ${site.url}`, err);
            return { siteUrl: site.url, siteLabel: site.label, libraries: [], error: 'Unable to load this site.' };
          }
        })
      );

      if (!cancelled) {
        setGroups(results);
        setLoading(false);
      }
    };

    if (props.sites.length === 0) {
      setGroups([]);
      setLoading(false);
    } else {
      loadLibraries().catch((err) => console.error(err));
    }

    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [props.sp, siteUrlsKey]);

  const toggleLibrary = (siteUrl: string, title: string): void => {
    const key = makeKey(siteUrl, title);
    const isExcluded = props.excludedLibraries.indexOf(key) !== -1;
    const next = isExcluded
      ? props.excludedLibraries.filter((k) => k !== key)
      : [...props.excludedLibraries, key];
    props.onChange(next);
  };

  return (
    <div className={styles.checklistWrapper}>
      <div className={styles.fieldLabel}>{props.label}</div>

      {props.sites.length === 0 ? (
        <p className={styles.hint}>Add at least one site above to choose which libraries to include.</p>
      ) : loading ? (
        <p className={styles.hint}>Loading libraries...</p>
      ) : (
        <div className={styles.groupsScroll}>
          {groups.map((group) => (
            <div key={group.siteUrl} className={styles.siteGroup}>
              <p className={styles.siteGroupTitle}>{group.siteLabel}</p>
              {group.error ? (
                <p className={styles.errorText}>{group.error}</p>
              ) : group.libraries.length === 0 ? (
                <p className={styles.hint}>No document libraries found on this site.</p>
              ) : (
                group.libraries.map((lib) => {
                  const key = makeKey(lib.siteUrl, lib.title);
                  const isChecked = props.excludedLibraries.indexOf(key) === -1;
                  return (
                    <label key={key} className={styles.checkboxRow}>
                      <input
                        type="checkbox"
                        checked={isChecked}
                        onChange={() => toggleLibrary(lib.siteUrl, lib.title)}
                      />
                      <span className={styles.checkboxLabel}>{lib.title}</span>
                      <span className={styles.checkboxCount}>{lib.itemCount}</span>
                    </label>
                  );
                })
              )}
            </div>
          ))}
        </div>
      )}
    </div>
  );
};

export default LibraryExclusionChecklist;
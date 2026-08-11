export interface ISiteEntry {
    url: string;
    label: string;
  }
  
  const KNOWN_PREFIXES = ['USAASC-PEOMS-PM-'];
  
  export const deriveLabel = (url: string): string => {
    const trimmed = url.trim().replace(/\/+$/, '');
    const segments = trimmed.split('/');
    const lastSegment = segments[segments.length - 1] || trimmed;
  
    for (const prefix of KNOWN_PREFIXES) {
      if (lastSegment.toUpperCase().indexOf(prefix.toUpperCase()) === 0) {
        const stripped = lastSegment.substring(prefix.length);
        return stripped.length > 0 ? stripped : lastSegment;
      }
    }
  
    return lastSegment;
  };
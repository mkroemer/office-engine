import { createContext, useContext } from 'react'

/**
 * The BCP 47 tag numbers typed into the views are read in (DocViewer's `locale`, canonical); undefined: a decimal
 * point and commas grouping by thousands. The host chooses it: the browser's language does not.
 */
export const LocaleContext = createContext<string | undefined>(undefined)

export const useLocale = () => useContext(LocaleContext)

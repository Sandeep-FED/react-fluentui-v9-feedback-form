import { Web } from "@pnp/sp/presets/all"

const Environment = {
  Site_URL: "https://nortonmcuk.sharepoint.com/sites/IntranetPortal/",
}
const Sp = Web(Environment.Site_URL)
export { Environment, Sp }

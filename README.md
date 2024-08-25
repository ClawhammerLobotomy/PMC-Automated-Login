This repo has been surplanted by https://github.com/ClawhammerLobotomy/PMC_Automation_Tools with up to date functionality.

[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)

# PMC Automated Login

Plex Manufacturing Cloud (PMC) Automated Login is a base to perform simple automated processes within the PMC environment.
Once the Plex class is called, you can perform any selenium based operations within the browser environment.

## Documentation

### Plex Class
    Plex(environment, user_id, password, company_code, pcn='', db='test', use_config=True, pcn_path=Path('pcn.json'))

- environment
  - Accepted options are Classic and UX. Determines how the program will log in
- user_id
  - Plex user ID
- password
  - Plex password
- company_code
  - Plex company code
- pcn
  - Optional. PCN number that would need to be selected after login.
  - Will not be needed if the account only has one PCN access or if using a UX login and operating in the account's main PCN.
- db
  - Optional. Defaults to 'test'. Accepted values are 'test' and 'prod'. Can be changed via the config file after it is created.
- use_config
  - Optional. Defaults to True. If false, it will bypass the configuration setup and use the supplied credentials.
  - Useful for GUI purposes and in order to not store plain text login details
- pcn_path
  - Optional. Allows for the main script to supply a path for the PCN json file.
  - Useful for compiled scripts where the end user wouldn't be able to supply the csv from the sql query.
  - Defaults to the working directory


## Contact

[sleemand@shapecorp.com](mailto:sleemand@shapecorp.com)

/****************************************************************************************
 *  TRU MANGOES — PROD DEPLOYMENT CONFIG  (temporary runner)
 *  --------------------------------------------------------------------------------------
 *  Code.gs ships with NEUTRAL defaults (no IDs, emails, inventory, or operator accounts).
 *  This file injects THIS deployment's real settings into Script Properties, so the code
 *  file itself stays generic and portable.
 *
 *  HOW TO USE (once per deployment):
 *    1. Paste Code.gs / JavaScript / Stylesheet / appsscript.json and SAVE.
 *    2. Paste THIS file into the editor (or paste applyProdConfig() into Code.gs temporarily).
 *    3. Run  applyProdConfig()  once.  Approve auth if prompted.
 *    4. Run  setMyOperatorPasswords()  once (edit the passwords below first).
 *    5. DELETE this file (and the password runner) so no secrets/values live in code.
 *    6. Deploy ▸ Manage deployments ▸ New version ▸ (access "Anyone") ▸ Deploy.
 *
 *  To view what's active later:  getConfig()    To wipe overrides:  clearConfig()
 ****************************************************************************************/

function applyProdConfig() {
  return setConfig({
    SPREADSHEET_ID: '1JHvRgLDvw0KUGubI73OhQjSaoPAq_XhMUv8VCFCxZMI',
    BUSINESS_NAME : 'TRU Mangoes',
    REQUIRE_ALLOWLIST: true,
    ALLOWLIST: [
      'mandalav@gmail.com',
      'opuslonestar@gmail.com',
      'Vinay.Yalamanchili@gmail.com'
    ],
    UI: {
      DEFAULT_TAB: 'orders',
      USER_ROLES: {
        'opuslonestar@gmail.com': 'admin',
        'mandalav@gmail.com'    : 'admin'
      },
      DEFAULT_ROLE: 'admin'
    },
    CANCEL_ENABLED: false,
    CANCEL_ADMINS: [
      'opuslonestar@gmail.com',
      'mandalav@gmail.com'
    ],
    INVENTORY: { banganapalli: 230, kesar: 0, rasalu: 80, himayat: 45 },
    OPERATORS: {
      ENABLED: true,
      ACCOUNTS: {
        'operator': { role: 'admin',  canCancel: true,  canEdit: true,  location: 'all'    },
        'frisco'  : { role: 'pickup', canCancel: false, canEdit: false, location: 'Frisco' },
        'plano'   : { role: 'pickup', canCancel: false, canEdit: false, location: 'Plano'  }
      }
    }
  });
}

/*  Set operator passwords (edit, run once, then DELETE this function).
 *  Usernames must match the ACCOUNTS keys above (case-insensitive).  */
function setMyOperatorPasswords() {
  setOperatorPassword('operator', 'IMPickup2026$');
  setOperatorPassword('frisco',   'frisco18');
  setOperatorPassword('plano',    'plano21');
}

/****************************************************************************************
 *  ALTERNATIVE — manual Script Property (instead of running applyProdConfig)
 *  Project Settings ▸ Script Properties ▸ Add property:
 *    Key:   config_overrides
 *    Value: (paste the single-line JSON below)
 *  --------------------------------------------------------------------------------------
 *  {"SPREADSHEET_ID":"1JHvRgLDvw0KUGubI73OhQjSaoPAq_XhMUv8VCFCxZMI","BUSINESS_NAME":"TRU Mangoes","REQUIRE_ALLOWLIST":true,"ALLOWLIST":["mandalav@gmail.com","opuslonestar@gmail.com","Vinay.Yalamanchili@gmail.com"],"UI":{"DEFAULT_TAB":"orders","USER_ROLES":{"opuslonestar@gmail.com":"admin","mandalav@gmail.com":"admin"},"DEFAULT_ROLE":"admin"},"CANCEL_ENABLED":false,"CANCEL_ADMINS":["opuslonestar@gmail.com","mandalav@gmail.com"],"INVENTORY":{"banganapalli":230,"kesar":0,"rasalu":80,"himayat":45},"OPERATORS":{"ENABLED":true,"ACCOUNTS":{"operator":{"role":"admin","canCancel":true,"canEdit":true,"location":"all"},"frisco":{"role":"pickup","canCancel":false,"canEdit":false,"location":"Frisco"},"plano":{"role":"pickup","canCancel":false,"canEdit":false,"location":"Plano"}}}}
 ****************************************************************************************/

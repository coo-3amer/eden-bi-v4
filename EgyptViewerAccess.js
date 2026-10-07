/**
 * EDEN BI — Egypt Viewer access policy
 * Role value: Egypt Viewer
 */

function egyptViewerRoleKey_(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[_-]+/g, ' ')
    .replace(/\s+/g, ' ');
}

function isEgyptViewerUser_(user) {
  const role = egyptViewerRoleKey_(user && user.role);
  return role === 'egypt viewer' || role === 'egypt transactions viewer';
}

function isEgyptNationality_(value) {
  const key = String(value || '').trim().toLowerCase();
  return [
    'egyptian',
    'egypt',
    'egyptian nationality',
    'مصري',
    'مصرى',
    'مصرية',
    'مصريه',
    'مصر'
  ].includes(key);
}

function assertEgyptViewerReadOnly_(user, actionName) {
  if (isEgyptViewerUser_(user)) {
    throw new Error(
      'Read-only account. Egypt Viewer is not allowed to ' +
      String(actionName || 'perform this action') + '.'
    );
  }
}

function assertEgyptViewerModuleBlocked_(user, moduleName) {
  if (isEgyptViewerUser_(user)) {
    throw new Error(
      'Access denied. Egypt Viewer can access Transactions / EDEN WALK and Brokers only. ' +
      String(moduleName || '') + ' is not available for this account.'
    );
  }
}

function egyptViewerAccessProfile_(user) {
  if (!isEgyptViewerUser_(user)) return null;

  return {
    role: 'Egypt Viewer',
    readOnly: true,
    canExport: true,
    exportFormats: ['CSV', 'Excel', 'PDF'],
    canWrite: false,
    allowedModules: ['transactions', 'brokers'],
    allowedProject: 'EDEN WALK',
    allowedNationality: 'Egyptian',
    allowedBranches: ['EGY - Al-Rehab', 'KSA - Jeddah', 'KSA - Riyadh']
  };
}

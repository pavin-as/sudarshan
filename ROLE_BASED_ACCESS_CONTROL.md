# Role-Based Access Control System

**This document is the main reference for access control in the Appointment Management System.**

- For implementation status and next steps, see `implementation-checklist.md`.
- For optimization details, see `optimization-activation-guide.md`.
- For analytics, see `RESCHEDULING_ANALYTICS_IMPLEMENTATION.md`.

---

## Overview

The appointment management system now includes a comprehensive role-based access control (RBAC) system with four hierarchical access levels. This system ensures that users can only access features and perform actions appropriate to their role and responsibility level.

## Access Levels

### Level 1: Director/Admin (Highest Privileges)
- **Roles**: `director`, `admin`
- **Description**: Full system access including user management, system settings, and all administrative functions
- **Permissions**:
  - All system functions
  - User management and role assignment
  - System settings and configuration
  - Delete appointments
  - Manage system settings
  - Access role management interface
  - Access visualization dashboard
  - Access settings configuration page

### Level 2: Doctor
- **Roles**: `doctor`
- **Description**: Medical and patient management, appointment archiving, and clinical operations
- **Permissions**:
  - Medical operations
  - Patient management
  - Appointment archiving
  - Clinical data access
  - Doctor-specific functions

### Level 3: Manager
- **Roles**: `manager`
- **Description**: Operational management, patient updates, analytics, and appointment management
- **Permissions**:
  - Operational management
  - Patient updates
  - Analytics and reporting
  - Appointment management
  - Reschedule appointments
  - Cancel appointments
  - Confirm appointments

### Level 4: General User (Lowest Privileges)
- **Roles**: `user`, `staff`, `receptionist`, `nurse`
- **Description**: Basic appointment operations including booking, viewing, and basic patient information
- **Permissions**:
  - Basic appointment operations
  - View patient information
  - Book appointments
  - View appointment lists

## Implementation Details

### Core Functions

#### `getUserAccessLevel(role)`
Converts a role string to its corresponding access level (1-4).

```javascript
function getUserAccessLevel(role) {
  const roleLevels = {
    'director': 1,
    'admin': 1,
    'doctor': 2,
    'manager': 3,
    'user': 4,
    'staff': 4,
    'receptionist': 4,
    'nurse': 4
  };
  
  const normalizedRole = role ? role.toLowerCase().trim() : 'user';
  return roleLevels[normalizedRole] || 4;
}
```

#### `hasAccessLevel(sessionToken, requiredLevel)`
Checks if a user has the minimum required access level.

```javascript
function hasAccessLevel(sessionToken, requiredLevel) {
  const session = getUserSession(sessionToken);
  if (!session) return false;
  
  const userLevel = getUserAccessLevel(session.role);
  return userLevel <= requiredLevel; // Lower number = higher access
}
```

### Level-Specific Functions

#### Level 1 Functions
- `isLevel1(sessionToken)` - Check if user is Level 1
- `hasAdminAccess(sessionToken)` - Check admin access
- `canDeleteAppointments(sessionToken)` - Can delete appointments
- `canManageSystemSettings(sessionToken)` - Can manage system settings
- `canAccessVisualization(sessionToken)` - Can access visualization dashboard
- `canAccessSettings(sessionToken)` - Can access settings configuration

#### Level 2 Functions
- `isLevel2(sessionToken)` - Check if user is Level 2
- `hasDoctorAccess(sessionToken)` - Check doctor access
- `canArchiveAppointments(sessionToken)` - Can archive appointments

#### Level 3 Functions
- `isLevel3(sessionToken)` - Check if user is Level 3
- `hasManagerAccess(sessionToken)` - Check manager access
- `canUpdatePatientDetails(sessionToken)` - Can update patient details
- `canViewAnalytics(sessionToken)` - Can view analytics
- `canRescheduleAppointments(sessionToken)` - Can reschedule appointments
- `canCancelAppointments(sessionToken)` - Can cancel appointments
- `canConfirmAppointments(sessionToken)` - Can confirm appointments

#### Level 4 Functions
- `isLevel4(sessionToken)` - Check if user is Level 4
- `canBookAppointments(sessionToken)` - Can book appointments

### Utility Functions

#### `getCurrentUserAccessLevel(sessionToken)`
Returns detailed information about the current user's access level.

```javascript
function getCurrentUserAccessLevel(sessionToken) {
  const session = getUserSession(sessionToken);
  if (!session) {
    return { success: false, message: "Invalid session" };
  }
  
  const level = getUserAccessLevel(session.role);
  const roleNames = {
    1: 'Director/Admin',
    2: 'Doctor', 
    3: 'Manager',
    4: 'General User'
  };
  
  return {
    success: true,
    level: level,
    roleName: roleNames[level] || 'Unknown',
    role: session.role
  };
}
```

#### `getAvailableRoles()`
Returns all available roles in the system with their levels and descriptions.

## Role Management Interface

### Access
The role management interface is accessible only to Level 1 users (Director/Admin) through:
- Dashboard quick actions (visible only to admins)
- Dashboard main menu (visible only to admins)

### Features
- **User Information Display**: Shows current user's role and access level
- **Access Level Summary**: Visual representation of all four access levels
- **Role Table**: Complete list of all roles with their permissions
- **Permission Checker**: Real-time display of current user's permissions

### Navigation
- **URL**: `?page=roleManagement&sessionToken=<token>`
- **Access Control**: Automatically redirects non-admin users to Dashboard

### Level 1 Restricted Pages
The following pages are accessible only to Level 1 users (Director/Admin):
- **Role Management**: `?page=roleManagement&sessionToken=<token>`
- **Visualization Dashboard**: `?page=Visualization&sessionToken=<token>`
- **Settings Configuration**: `?page=settings&sessionToken=<token>`

All these pages will automatically redirect non-admin users to the Dashboard.

## Integration with Existing Functions

### Updated Functions
The following existing functions have been updated to use the new RBAC system:

#### `isAdmin(sessionToken)`
Now uses `hasAdminAccess(sessionToken)` internally.

#### `hasAccessToPatientMaster(sessionToken)`
Now uses `canUpdatePatientDetails(sessionToken)` internally, allowing Level 1, 2, and 3 users to access patient master.

### Function Access Control
All sensitive functions now include access level checks:

```javascript
function deleteAppointment(appointmentId, sessionToken) {
  // Check if user can delete appointments (Level 1 only)
  if (!canDeleteAppointments(sessionToken)) {
    return { success: false, message: "Unauthorized: Admin access required" };
  }
  // ... rest of function
}
```

## Database Structure

### Login Sheet
The system expects the following structure in the Login sheet:
- **Column A**: Username
- **Column B**: Password
- **Column C**: Role (one of: director, admin, doctor, manager, user, staff, receptionist, nurse)
- **Column D**: Additional user data (optional)
- **Column E**: Email (optional)

### Role Assignment
Roles are case-insensitive and automatically normalized. Unknown roles default to Level 4 (General User).

## Security Features

### Session Validation
- All access checks validate session tokens
- Invalid or expired sessions are automatically rejected
- Session timeout after 24 hours

### Access Logging
- All access control decisions are logged
- Failed access attempts are tracked
- Session creation and deletion are logged

### Error Handling
- Graceful degradation for missing roles
- Default access level for unknown roles
- Comprehensive error messages for debugging

## Usage Examples

### Frontend Access Control
```javascript
// Check if user can access a feature
google.script.run
  .withSuccessHandler(function(hasAccess) {
    if (hasAccess) {
      showFeature();
    } else {
      showAccessDenied();
    }
  })
  .canViewAnalytics(sessionToken);
```

### Backend Function Protection
```javascript
function sensitiveFunction(sessionToken, data) {
  // Check access level before proceeding
  if (!hasAccessLevel(sessionToken, 2)) {
    return { success: false, message: "Insufficient privileges" };
  }
  
  // Proceed with function logic
  // ...
}
```

### UI Element Visibility
```javascript
// Show/hide elements based on access level
function updateUIBasedOnAccess() {
  google.script.run
    .withSuccessHandler(function(hasAccess) {
      document.getElementById('adminPanel').style.display = 
        hasAccess ? 'block' : 'none';
    })
    .hasAdminAccess(sessionToken);
}
```

## Migration Guide

### For Existing Users
1. **Role Assignment**: Ensure all users have appropriate roles in the Login sheet
2. **Access Testing**: Test each user's access to ensure proper functionality
3. **UI Updates**: Update any custom UI elements to use the new access control functions

### For New Features
1. **Access Level Planning**: Determine the minimum access level required for the feature
2. **Function Implementation**: Use appropriate access check functions
3. **UI Integration**: Update UI to show/hide elements based on access level
4. **Testing**: Test with users of different access levels

## Troubleshooting

### Common Issues

#### "Access Denied" Errors
- Check user's role in Login sheet
- Verify session token is valid
- Ensure function is calling correct access check

#### Role Management Not Visible
- Verify user has admin role
- Check browser console for JavaScript errors
- Ensure roleManagement.html file exists

#### Unexpected Access Levels
- Check role spelling in Login sheet (case-insensitive)
- Verify getUserAccessLevel function is working
- Check session data for correct role assignment

### Debug Functions
```javascript
// Get current user's access information
google.script.run
  .withSuccessHandler(function(info) {
    console.log('User access info:', info);
  })
  .getCurrentUserAccessLevel(sessionToken);

// Check specific permission
google.script.run
  .withSuccessHandler(function(hasPermission) {
    console.log('Has permission:', hasPermission);
  })
  .canDeleteAppointments(sessionToken);
```

## Future Enhancements

### Planned Features
- **Role Hierarchy Management**: Allow custom role hierarchies
- **Permission Granularity**: More fine-grained permissions
- **Access Audit Trail**: Detailed logging of all access attempts
- **Role Templates**: Predefined role templates for common scenarios
- **Bulk Role Management**: Manage multiple users at once

### Extension Points
- **Custom Roles**: Add new roles with custom access levels
- **Permission Overrides**: Temporary permission grants
- **Role Inheritance**: Hierarchical role relationships
- **Time-based Access**: Access restrictions based on time/date

## Support

For issues or questions regarding the role-based access control system:
1. Check this documentation first
2. Review the console logs for error messages
3. Verify user roles in the Login sheet
4. Test with different user accounts
5. Contact system administrator for role assignments 
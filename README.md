# Appointment Management System

## Overview

This project is a comprehensive Appointment Management System designed for hospitals and clinics. It provides robust features for appointment booking, rescheduling analytics, role-based access control, and performance optimizations.

## Key Features

- **Appointment Booking & Management**: Book, view, reschedule, and cancel appointments with advanced search and filtering.
- **Rescheduling Analytics**: Track and analyze appointment rescheduling patterns, staff accountability, and high-risk appointments. (See `RESCHEDULING_ANALYTICS_IMPLEMENTATION.md`)
- **Role-Based Access Control (RBAC)**: Four-level access system (Admin, Doctor, Manager, User/Staff) to ensure secure and appropriate feature access. (See `ROLE_BASED_ACCESS_CONTROL.md`)
- **Performance Optimizations**: Fast search, caching, batch operations, and a modern responsive UI. (See `optimization-activation-guide.md`)

## Documentation

- [Implementation Checklist](./implementation-checklist.md): Status, roadmap, and next steps
- [Optimization Guide](./optimization-activation-guide.md): Technical details of performance improvements
- [Rescheduling Analytics](./RESCHEDULING_ANALYTICS_IMPLEMENTATION.md): Analytics features and usage
- [Role-Based Access Control](./ROLE_BASED_ACCESS_CONTROL.md): Access levels and permissions

## Setup & Usage

1. **Google Apps Script Backend**: Deploy `code.js` as a Google Apps Script bound to your Google Sheet.
2. **Frontend**: Use the provided HTML files for the web interface (upload as Apps Script HTML files or host as needed).
3. **Configuration**: Ensure your Google Sheet has the required sheets (e.g., `appointment`, `patientMaster`, `Login`, etc.) and columns as described in the documentation.
4. **User Roles**: Assign user roles in the `Login` sheet for proper access control.

## Contributing

Contributions are welcome! Please:
- Review the documentation before submitting changes
- Ensure new features include appropriate access control and analytics integration
- Update documentation as needed

## Support

For questions or issues:
- Review the relevant documentation files
- Check the Google Apps Script logs for errors
- Contact the system administrator for access or configuration help

---

**For full technical and feature documentation, see the referenced .md files in this repository.** 
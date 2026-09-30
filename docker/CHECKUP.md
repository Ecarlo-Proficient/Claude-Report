# Synology check-up - what to look at on the box

Step 0 of the setup (`README.md`): a list of what to check before the office server goes on. Fix what needs fixing.

Remote workers reach the Synology through the **Fortinet VPN** - that is the approved way in and it stays.

1. **The box**
   - [ ] Model, DSM version, date of the last update
   - [ ] Are updates set to install automatically? (yes / no / notify only)
   - [ ] Disk and storage health: any warnings?
   - [ ] Installed apps / packages (list)
   - [ ] What time does the Synology show vs a phone? Is it set to sync its clock automatically (NTP)? Which time zone?
   - [ ] Is it on a battery backup (UPS)? Does it shut down cleanly on a power cut?

2. **Who can get in**
   - [ ] Every user account and whether it is an admin
   - [ ] Is the built-in "admin" account on?
   - [ ] Two-step login on any account? Which?
   - [ ] Guest account on?
   - [ ] Auto-block for wrong passwords on?

3. **How people reach it from outside**
   - [ ] QuickConnect on?
   - [ ] DDNS set up? (the address, if yes)
   - [ ] On the Fortinet: any port forwards / virtual IPs pointing at the Synology? (list)
   - [ ] On the Fortinet: what address range do VPN users get?
   - [ ] Does the Fortinet VPN ask for a second step (code / app) at login?
   - [ ] From outside the office with the VPN OFF (phone hotspot): can you reach the Synology login page or files?

4. **What it is serving**
   - [ ] Services on/off: SMB, AFP, FTP, WebDAV, SSH, Telnet, SNMP, others
   - [ ] The Synology's own firewall on? Its rules?
   - [ ] DSM login page on http, https, or both?

5. **Shared folders**
   - [ ] Every shared folder and which groups / users can open it
         (do NOT open or list anything inside `Proinfo/Items` - only who has access to it)
   - [ ] Any folder open to "everyone" or guest?

6. **Backups**
   - [ ] Snapshots on? Which folders, how often, kept how long?
   - [ ] Can an admin delete those snapshots, or are they locked?
   - [ ] Any backup to somewhere OFF the Synology? Where, how often, last success?
   - [ ] Has anyone ever restored from it?

7. **Already running**
   - [ ] Container Manager installed? Any containers (names)?
   - [ ] Scheduled tasks in Task Scheduler (list)
   - [ ] Anything syncing to the cloud (Cloud Sync, Synology Drive, Hyper Backup)? What and where?

8. **Warnings already on file**
   - [ ] Run Security Advisor (it only scans) and send the results page
   - [ ] Log Center: failed logins, logins from unknown places?

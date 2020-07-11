# Notes for build: {{buildDetails.buildNumber}}

**Build Number**: {{buildDetails.id}}

**Build Trigger PR Number**: {{lookup buildDetails.triggerInfo 'pr.number'}}

**Build Source Branch** - {{buildDetails.sourceBranch}}

**Build Name** - {{buildDetails.definition.name}}

**Build Reason** - {{buildDetails.reason}}

# Associated Pull Requests ({{pullRequests.length}})
{{#forEach pullRequests}}
{{#if isFirst}}### Associated Pull Requests (only shown if  PR) {{/if}}
*  **PR {{this.id}}**  {{this.title}}
{{/forEach}}

# Global list of WI ({{workItems.length}})
{{#forEach workItems}}
{{#if isFirst}}## Associated Work Items (only shown if  WI) {{/if}}
*  **{{this.id}}**  {{lookup this.fields 'System.Title'}}
  - **WIT** {{lookup this.fields 'System.WorkItemType'}}
  - **Tags** {{lookup this.fields 'System.Tags'}}
  - **Assigned** {{#with (lookup this.fields 'System.AssignedTo')}} {{displayName}} {{/with}}
{{/forEach}}

# Global list of CS ({{commits.length}})
{{#forEach commits}}
{{#if isFirst}}### Associated commits{{/if}}
* ** ID{{this.id}}**
  -  **Message:** {{this.message}}
  -  **Commited by:** {{this.author.displayName}}
  -  **Timestamp:** {{this.timestamp}}
  -  **FileCount:** {{this.changes.length}}
{{#forEach this.changes}}
      -  **File path (TFVC or TfsGit):** {{this.item.path}}
{{/forEach}}
{{/forEach}}

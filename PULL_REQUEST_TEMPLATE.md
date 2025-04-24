## PR Title
[Type]: Brief Description  
**Example**:  
[FEATURE] Support new Mode
[BUGFIX] "Connect me to another external party"

## Description  
Explain your changes in detail. Include:  
- Purpose of changes  
- Technical approach used  
- Screenshots/GIFs if applicable  

**Example**:  
"Added support for ?:  
- Implemented 
- Added 
- Included test projects in /samples folder  
- Updated garbage logs to show full paths (see test.log)"  

## Proposed Changes  
- [ ] 🐛 Bug fix (include issue #)  
- [ ] ✨ Feature implementation (link to discussion #)  
- [ ] 📚 Documentation update  
- [ ] 🧹 Code cleanup/refactor  

## Checklist  
- [ ] Tested with TwinBASIC/VB6 IDE (version 6.0 SP6)  
- [ ] Verified on Windows 10/11  
- [ ] Updated README.md if applicable  
- [ ] Added/updated garbage log samples  
- [ ] No debug code/console logs left  
- [ ] Follows [coding standards](https://github.com/aidajulia/vb6-comment-cleaner-pro/wiki/Coding-Standards)  

## Related Issues  
- Closes #123  
- Fixes #45  
- Related to #67  

## Testing Summary  
| Test Case | Result |  
|----------|--------|  
| Empty .cls file | ✅ Pass |  
| Nested folders | ✅ Pass |  
| UTF-8 comments | ⚠️ Needs review |  

## Additional Notes  
- Breaking changes: [Yes/No]  
- Dependencies: [List any new dependencies]  
- Performance impact: [CPU/Memory metrics]  

---

**Preview of Changes**  
```vb
' Before
' Sub ProcessFiles()
'   ' Old single-folder logic

' After
Sub ProcessFiles(Optional ByVal ProcessSubfolders As Boolean = False)
    ' New recursive logic

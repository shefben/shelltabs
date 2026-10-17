#include "GroupStore.h"
#include "Utilities.h"

#include <iostream>

namespace {
// Substitute only the storage boundary; exercise the real group parser and writer.
std::wstring contents = L"version|2\ngroup|Work|#0078D7|solid\ntab|C:\\Original\n"
                        L"group|Other|#0078D7|solid\ntab|C:\\Other\n";
bool failWrites = false;
int writes = 0;
}

namespace shelltabs {
std::wstring GetShellTabsDataDirectory() {
    wchar_t path[MAX_PATH]{};
    GetTempPathW(MAX_PATH, path);
    return path;
}
bool ReadUtf8File(const std::wstring&, std::wstring* text, bool* exists) {
    *text = contents;
    if (exists) {
        *exists = true;
    }
    return true;
}
bool WriteUtf8File(const std::wstring&, std::wstring_view text) {
    ++writes;
    if (failWrites) {
        return false;
    }
    contents = text;
    return true;
}
}

int main() {
    auto& store = shelltabs::GroupStore::Instance();
    bool success = true;
    auto expect = [&](bool value, const char* message) {
        if (!value) {
            std::cerr << message << '\n';
            success = false;
        }
    };
    expect(store.Load(), "Existing saved groups must load");
    expect(writes == 0, "Loading a group must not rewrite its saved folders");
    expect(store.AddTabPath(L"Work", L"C:\\Added"), "Explicit save must append the folder");
    const auto* work = store.Find(L"Work");
    expect(work && work->tabPaths == std::vector<std::wstring>{L"C:\\Original", L"C:\\Added"},
           "Saving one folder must preserve previously saved paths and order");
    expect(contents.find(L"tab|C:\\Added\n") != std::wstring::npos,
           "Explicitly added folder must be serialized");
    const int previousWrites = writes;
    expect(store.AddTabPath(L"work", L"c:/added/"), "Equivalent Windows folder must be accepted");
    expect(writes == previousWrites && work->tabPaths.size() == 2,
           "Equivalent Windows paths must not create duplicates or rewrite storage");
    expect(!store.AddTabPath(L"Missing", L"C:\\Temporary"), "Saving must not recreate deleted groups");
    expect(!store.AddTabPath(L"Work", L""), "Empty folders must not be saved");
    expect(store.Find(L"Other")->tabPaths == std::vector<std::wstring>{L"C:\\Other"},
           "Saving must not change another group's folders");
    failWrites = true;
    const auto previousContents = contents;
    expect(!store.AddTabPath(L"Work", L"C:\\Failed"), "Failed writes must be reported");
    expect(work->tabPaths.size() == 2 && contents == previousContents,
           "Failed saves must leave both memory and persisted folders unchanged");
    shelltabs::SavedGroup newGroup;
    newGroup.name = L"New island";
    newGroup.tabPaths = {L"C:\\ClickedFolder"};
    expect(!store.Upsert(newGroup), "Creating a group must report storage failure");
    expect(!store.Find(newGroup.name), "Failed creation must not leave an unsaved group in memory");
    // Inserting a definition can reallocate the store's vector, even on rollback.
    shelltabs::SavedGroup replacement = *store.Find(L"Work");
    replacement.tabPaths = {L"C:\\Replacement"};
    expect(!store.Upsert(replacement), "Failed definition updates must be reported");
    expect(store.Find(L"Work")->tabPaths == std::vector<std::wstring>{L"C:\\Original", L"C:\\Added"},
           "Failed definition updates must preserve original saved folders");
    failWrites = false;
    expect(store.Upsert(newGroup), "An unsaved island must be promotable to a saved group");
    expect(store.Find(newGroup.name)->tabPaths == newGroup.tabPaths,
           "Promoting an island must save only the explicitly chosen folder");
    return success ? 0 : 1;
}

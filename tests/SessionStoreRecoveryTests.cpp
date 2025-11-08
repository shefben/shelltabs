#include "SessionStore.h"

#include <iostream>
#include <string>
#include <vector>

namespace {

using shelltabs::SessionStore;

SessionStore::RecoverableSessionCandidate MakeCandidate(const std::wstring& token, int markers,
                                                         uint64_t ticks) {
    SessionStore::RecoverableSessionCandidate candidate;
    candidate.token = token;
    candidate.storagePath = L"session-" + token + L".db";
    candidate.lastActivityTicks = ticks;
    if (markers >= 1) {
        candidate.hasLock = true;
    }
    if (markers >= 2) {
        candidate.hasTemp = true;
    }
    if (markers >= 3) {
        candidate.hasCheckpoint = true;
    }
    return candidate;
}

bool TestPrefersMoreMarkers() {
    std::vector<SessionStore::RecoverableSessionCandidate> candidates;
    candidates.push_back(MakeCandidate(L"lock-only", 1, 10));
    candidates.push_back(MakeCandidate(L"lock-temp", 2, 20));
    candidates.push_back(MakeCandidate(L"lock-temp-prev", 3, 5));

    const auto selected = SessionStore::SelectRecoverableSession(candidates);
    if (!selected || selected->token != L"lock-temp-prev") {
        std::wcerr << L"[TestPrefersMoreMarkers] Expected lock-temp-prev but received "
                   << (selected ? selected->token : L"<none>") << std::endl;
        return false;
    }
    return true;
}

bool TestPrefersNewestWhenMarkersEqual() {
    std::vector<SessionStore::RecoverableSessionCandidate> candidates;
    candidates.push_back(MakeCandidate(L"older", 1, 10));
    candidates.push_back(MakeCandidate(L"newer", 1, 50));

    const auto selected = SessionStore::SelectRecoverableSession(candidates);
    if (!selected || selected->token != L"newer") {
        std::wcerr << L"[TestPrefersNewestWhenMarkersEqual] Expected newer but received "
                   << (selected ? selected->token : L"<none>") << std::endl;
        return false;
    }
    return true;
}

bool TestReturnsNullWhenNoMarkers() {
    std::vector<SessionStore::RecoverableSessionCandidate> candidates;
    candidates.push_back(MakeCandidate(L"no-markers", 0, 100));

    const auto selected = SessionStore::SelectRecoverableSession(candidates);
    if (selected) {
        std::wcerr << L"[TestReturnsNullWhenNoMarkers] Expected no selection but received "
                   << selected->token << std::endl;
        return false;
    }
    return true;
}

}  // namespace

int wmain() {
    struct TestCase {
        const wchar_t* name;
        bool (*fn)();
    } tests[] = {
        {L"TestPrefersMoreMarkers", &TestPrefersMoreMarkers},
        {L"TestPrefersNewestWhenMarkersEqual", &TestPrefersNewestWhenMarkersEqual},
        {L"TestReturnsNullWhenNoMarkers", &TestReturnsNullWhenNoMarkers},
    };

    bool success = true;
    for (const auto& test : tests) {
        if (!test.fn()) {
            success = false;
        }
    }

    return success ? 0 : 1;
}

#pragma once

namespace shelltabs {

// Acquires a reference to the global MinHook library. The library is
// initialized on the first successful acquisition and reference counted for
// subsequent users. The optional context string is used for logging.
bool AcquireMinHook(const wchar_t* context) noexcept;

// Releases a previously acquired MinHook reference. When the last reference is
// released, the MinHook library is uninitialized. The optional context string
// is used for logging.
void ReleaseMinHook(const wchar_t* context) noexcept;

// Helper that acquires MinHook on construction and automatically releases it on
// destruction unless dismissed. This is primarily intended for use within
// initialization routines so that early returns unwind correctly.
class MinHookScopedAcquire {
public:
    explicit MinHookScopedAcquire(const wchar_t* context) noexcept;
    MinHookScopedAcquire(const MinHookScopedAcquire&) = delete;
    MinHookScopedAcquire& operator=(const MinHookScopedAcquire&) = delete;
    MinHookScopedAcquire(MinHookScopedAcquire&&) noexcept = delete;
    MinHookScopedAcquire& operator=(MinHookScopedAcquire&&) noexcept = delete;
    ~MinHookScopedAcquire();

    [[nodiscard]] bool IsAcquired() const noexcept { return m_acquired; }

    // Releases the held reference immediately if one is present.
    void Release() noexcept;

    // Prevents the destructor from releasing the reference. Use this when the
    // caller wants to keep MinHook initialized beyond the current scope (for
    // example, when initialization succeeded and shutdown code will release it
    // later).
    void Dismiss() noexcept { m_acquired = false; }

private:
    const wchar_t* m_context;
    bool m_acquired;
};

}  // namespace shelltabs


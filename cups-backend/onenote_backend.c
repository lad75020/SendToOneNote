#include <cups/cups.h>
#include <cups/backend.h>
#include <stdio.h>
#include <stdlib.h>
#include <unistd.h>
#include <string.h>
#include <errno.h>
#include <fcntl.h>
#include <pwd.h>
#include <limits.h>
#include <sys/stat.h>
#include <time.h>

static int file_starts_with(const char *path, const char *prefix) {
    int fd = open(path, O_RDONLY);
    if (fd < 0) return 0;
    char buf[16];
    ssize_t n = read(fd, buf, sizeof(buf));
    close(fd);
    if (n <= 0) return 0;
    size_t plen = strlen(prefix);
    if ((size_t)n < plen) return 0;
    return (memcmp(buf, prefix, plen) == 0);
}

static int ensure_dir(const char *path, mode_t mode) {
    if (!path || !path[0]) return -1;
    if (mkdir(path, mode) == 0) return 0;
    if (errno == EEXIST) return 0;
    return -1;
}

static int ensure_dirs_recursive(const char *path, mode_t mode) {
    // Minimal mkdir -p for absolute paths.
    char tmp[PATH_MAX];
    if (!path || !path[0] || strlen(path) >= sizeof(tmp)) return -1;
    strncpy(tmp, path, sizeof(tmp));
    tmp[sizeof(tmp) - 1] = '\0';

    for (char *p = tmp + 1; *p; p++) {
        if (*p == '/') {
            *p = '\0';
            if (ensure_dir(tmp, mode) != 0) {
                if (errno != EEXIST) return -1;
            }
            *p = '/';
        }
    }
    return ensure_dir(tmp, mode);
}

static int copy_file(const char *src, const char *dst, mode_t mode) {
    int in = open(src, O_RDONLY);
    if (in < 0) return -1;

    int out = open(dst, O_WRONLY | O_CREAT | O_TRUNC, mode);
    if (out < 0) {
        close(in);
        return -1;
    }

    char buf[8192];
    ssize_t n;
    while ((n = read(in, buf, sizeof(buf))) > 0) {
        ssize_t off = 0;
        while (off < n) {
            ssize_t w = write(out, buf + off, (size_t)(n - off));
            if (w < 0) {
                close(in);
                close(out);
                return -1;
            }
            off += w;
        }
    }

    close(in);
    close(out);
    return (n < 0) ? -1 : 0;
}

int main(int argc, char *argv[]) {
    // Discovery mode: no arguments.
    if (argc == 1) {
        return CUPS_BACKEND_OK;
    }

    if (argc != 6 && argc != 7) {
        fprintf(stderr, "DEBUG: invalid argc=%d\n", argc);
        return CUPS_BACKEND_FAILED;
    }

    const char *job_id  = argv[1];
    const char *user    = argv[2];
    const char *title   = argv[3];
    const char *copies  = argv[4];
    const char *options = argv[5];
    const char *file    = (argc == 7) ? argv[6] : NULL; // else read stdin

    (void)copies;
    (void)options;

    const char *tmpdir = getenv("TMPDIR");
    if (!tmpdir || !tmpdir[0]) tmpdir = "/private/var/spool/cups/tmp";

    // Write incoming job data to a temp file (no assumption about format).
    char templatePath[1024];
    snprintf(templatePath, sizeof(templatePath), "%s/onenote-print-XXXXXX", tmpdir);

    int outfd = mkstemp(templatePath);
    if (outfd < 0) {
        fprintf(stderr, "onenote backend: failed to create temp file '%s': %s\n", templatePath, strerror(errno));
        return CUPS_BACKEND_FAILED;
    }

    int infd = STDIN_FILENO;
    if (file && file[0]) {
        infd = open(file, O_RDONLY);
        if (infd < 0) {
            close(outfd);
            fprintf(stderr, "onenote backend: failed to open input file '%s': %s\n", file, strerror(errno));
            return CUPS_BACKEND_FAILED;
        }
    }

    size_t total_written = 0;
    char buffer[8192];
    ssize_t n;
    while ((n = read(infd, buffer, sizeof(buffer))) > 0) {
        ssize_t off = 0;
        while (off < n) {
            ssize_t wn = write(outfd, buffer + off, (size_t)(n - off));
            if (wn < 0) {
                if (infd != STDIN_FILENO) close(infd);
                close(outfd);
                fprintf(stderr, "onenote backend: failed to write temp file: %s\n", strerror(errno));
                return CUPS_BACKEND_FAILED;
            }
            off += wn;
            total_written += (size_t)wn;
        }
    }

    if (n < 0) {
        if (infd != STDIN_FILENO) close(infd);
        close(outfd);
        fprintf(stderr, "onenote backend: failed to read input data: %s\n", strerror(errno));
        return CUPS_BACKEND_FAILED;
    }

    if (infd != STDIN_FILENO) close(infd);
    close(outfd);

    fprintf(stderr, "onenote backend: wrote %zu bytes to %s\n", total_written, templatePath);

    // Decide output extension: keep PS as PS; keep PDF as PDF; default to PDF.
    const char *ext = "pdf";
    if (file_starts_with(templatePath, "%!PS")) {
        ext = "ps";
        fprintf(stderr, "onenote backend: detected PostScript; queueing as .ps for helper conversion\n");
    } else if (file_starts_with(templatePath, "%PDF")) {
        ext = "pdf";
    }

    struct passwd *pw = getpwnam(user);
    if (!pw || !pw->pw_dir) {
        fprintf(stderr, "onenote backend: getpwnam(%s) failed\n", user);
        return CUPS_BACKEND_FAILED;
    }

    const char *rootDir = "/Users/Shared/OneNoteHelper";
    const char *incomingDir = "/Users/Shared/OneNoteHelper/Incoming";
    const char *processingDir = "/Users/Shared/OneNoteHelper/Processing";
    const char *doneDir = "/Users/Shared/OneNoteHelper/Done";
    const char *failedDir = "/Users/Shared/OneNoteHelper/Failed";

    const char *dirs[] = { rootDir, incomingDir, processingDir, doneDir, failedDir };
    for (size_t i = 0; i < sizeof(dirs)/sizeof(dirs[0]); i++) {
        if (ensure_dirs_recursive(dirs[i], 0777) != 0) {
            fprintf(stderr, "onenote backend: failed to create dir '%s': %s\n", dirs[i], strerror(errno));
            return CUPS_BACKEND_FAILED;
        }
        (void)chmod(dirs[i], 01777);
    }

    time_t now = time(NULL);
    char baseName[256];
    snprintf(baseName, sizeof(baseName), "job-%s-%ld", job_id, (long)now);

    char destDoc[PATH_MAX];
    char destJson[PATH_MAX];
    snprintf(destDoc, sizeof(destDoc), "%s/%s.%s", incomingDir, baseName, ext);
    snprintf(destJson, sizeof(destJson), "%s/%s.json", incomingDir, baseName);

    // Move/copy document into Incoming
    if (rename(templatePath, destDoc) != 0) {
        if (copy_file(templatePath, destDoc, 0644) != 0) {
            fprintf(stderr, "onenote backend: failed to move/copy '%s' -> '%s': %s\n", templatePath, destDoc, strerror(errno));
            return CUPS_BACKEND_FAILED;
        }
        (void)unlink(templatePath);
    }

    if (chown(destDoc, pw->pw_uid, pw->pw_gid) != 0) {
        fprintf(stderr, "onenote backend: chown failed for '%s': %s (continuing)\n", destDoc, strerror(errno));
    }
    (void)chmod(destDoc, 0644);

    // Minimal JSON (escape only quotes/backslashes/newlines in title)
    char escTitle[2048];
    size_t ei = 0;
    for (size_t i = 0; title[i] != '\0' && ei < sizeof(escTitle) - 1; ++i) {
        char c = title[i];
        if (c == '"' || c == '\\') {
            if (ei + 2 >= sizeof(escTitle)) break;
            escTitle[ei++] = '\\';
            escTitle[ei++] = c;
        } else if (c == '\n' || c == '\r' || c == '\t') {
            if (ei + 2 >= sizeof(escTitle)) break;
            escTitle[ei++] = '\\';
            escTitle[ei++] = (c == '\n') ? 'n' : (c == '\r') ? 'r' : 't';
        } else {
            escTitle[ei++] = c;
        }
    }
    escTitle[ei] = '\0';

    FILE *jf = fopen(destJson, "w");
    if (!jf) {
        fprintf(stderr, "onenote backend: failed to write json '%s': %s\n", destJson, strerror(errno));
        return CUPS_BACKEND_FAILED;
    }
    fprintf(jf,
            "{\n"
            "  \"file\": \"%s\",\n"
            "  \"title\": \"%s\",\n"
            "  \"user\": \"%s\",\n"
            "  \"job\": \"%s\"\n"
            "}\n",
            destDoc, escTitle, user, job_id);
    fclose(jf);

    if (chown(destJson, pw->pw_uid, pw->pw_gid) != 0) {
        fprintf(stderr, "onenote backend: chown failed for '%s': %s (continuing)\n", destJson, strerror(errno));
    }
    (void)chmod(destJson, 0644);

    fprintf(stderr, "onenote backend: queued for helper: %s (+json)\n", destDoc);

    return CUPS_BACKEND_OK;
}

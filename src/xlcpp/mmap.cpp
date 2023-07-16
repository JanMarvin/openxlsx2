#include "mmap.h"

#ifndef _WIN32
#include <sys/mman.h>
#include <sys/types.h>
#include <sys/stat.h>
#include <errno.h>
#endif

using namespace std;

#ifndef _WIN32
errno_error::errno_error(string_view function, int en) : msg(function) {
	msg += " failed (";

	switch (en) {
		case EPERM:
			msg += "EPERM";
			break;
		case ENOENT:
			msg += "ENOENT";
			break;
		case ESRCH:
			msg += "ESRCH";
			break;
		case EINTR:
			msg += "EINTR";
			break;
		case EIO:
			msg += "EIO";
			break;
		case ENXIO:
			msg += "ENXIO";
			break;
		case E2BIG:
			msg += "E2BIG";
			break;
		case ENOEXEC:
			msg += "ENOEXEC";
			break;
		case EBADF:
			msg += "EBADF";
			break;
		case ECHILD:
			msg += "ECHILD";
			break;
		case EAGAIN:
			msg += "EAGAIN";
			break;
		case ENOMEM:
			msg += "ENOMEM";
			break;
		case EACCES:
			msg += "EACCES";
			break;
		case EFAULT:
			msg += "EFAULT";
			break;
		case ENOTBLK:
			msg += "ENOTBLK";
			break;
		case EBUSY:
			msg += "EBUSY";
			break;
		case EEXIST:
			msg += "EEXIST";
			break;
		case EXDEV:
			msg += "EXDEV";
			break;
		case ENODEV:
			msg += "ENODEV";
			break;
		case ENOTDIR:
			msg += "ENOTDIR";
			break;
		case EISDIR:
			msg += "EISDIR";
			break;
		case EINVAL:
			msg += "EINVAL";
			break;
		case ENFILE:
			msg += "ENFILE";
			break;
		case EMFILE:
			msg += "EMFILE";
			break;
		case ENOTTY:
			msg += "ENOTTY";
			break;
		case ETXTBSY:
			msg += "ETXTBSY";
			break;
		case EFBIG:
			msg += "EFBIG";
			break;
		case ENOSPC:
			msg += "ENOSPC";
			break;
		case ESPIPE:
			msg += "ESPIPE";
			break;
		case EROFS:
			msg += "EROFS";
			break;
		case EMLINK:
			msg += "EMLINK";
			break;
		case EPIPE:
			msg += "EPIPE";
			break;
		case EDOM:
			msg += "EDOM";
			break;
		case ERANGE:
			msg += "ERANGE";
			break;
		case ENAMETOOLONG:
			msg += "ENAMETOOLONG";
			break;
		case ENOLCK:
			msg += "ENOLCK";
			break;
		case ENOSYS:
			msg += "ENOSYS";
			break;
		case ENOTEMPTY:
			msg += "ENOTEMPTY";
			break;
		case ELOOP:
			msg += "ELOOP";
			break;
		case ENOMSG:
			msg += "ENOMSG";
			break;
		case EIDRM:
			msg += "EIDRM";
			break;
		// case ECHRNG:
		// 	msg += "ECHRNG";
		// 	break;
		// case EL2NSYNC:
		// 	msg += "EL2NSYNC";
		// 	break;
		// case EL3HLT:
		// 	msg += "EL3HLT";
		// 	break;
		// case EL3RST:
		// 	msg += "EL3RST";
		// 	break;
		// case ELNRNG:
		// 	msg += "ELNRNG";
		// 	break;
		// case EUNATCH:
		// 	msg += "EUNATCH";
		// 	break;
		// case ENOCSI:
		// 	msg += "ENOCSI";
		// 	break;
		// case EL2HLT:
		// 	msg += "EL2HLT";
		// 	break;
		// case EBADE:
		// 	msg += "EBADE";
		// 	break;
		// case EBADR:
		// 	msg += "EBADR";
		// 	break;
		// case EXFULL:
		// 	msg += "EXFULL";
		// 	break;
		// case ENOANO:
		// 	msg += "ENOANO";
		// 	break;
		// case EBADRQC:
		// 	msg += "EBADRQC";
		// 	break;
		// case EBADSLT:
		// 	msg += "EBADSLT";
		// 	break;
		// case EDEADLOCK:
		// 	msg += "EDEADLOCK";
		// 	break;
		// case EBFONT:
		// 	msg += "EBFONT";
		// 	break;
		case ENOSTR:
			msg += "ENOSTR";
			break;
		case ENODATA:
			msg += "ENODATA";
			break;
		case ETIME:
			msg += "ETIME";
			break;
		case ENOSR:
			msg += "ENOSR";
			break;
		// case ENONET:
		// 	msg += "ENONET";
		// 	break;
		// case ENOPKG:
		// 	msg += "ENOPKG";
		// 	break;
		case EREMOTE:
			msg += "EREMOTE";
			break;
		case ENOLINK:
			msg += "ENOLINK";
			break;
		// case EADV:
		// 	msg += "EADV";
		// 	break;
		// case ESRMNT:
		// 	msg += "ESRMNT";
		// 	break;
		// case ECOMM:
		// 	msg += "ECOMM";
		// 	break;
		case EPROTO:
			msg += "EPROTO";
			break;
		case EMULTIHOP:
			msg += "EMULTIHOP";
			break;
		// case EDOTDOT:
		// 	msg += "EDOTDOT";
		// 	break;
		case EBADMSG:
			msg += "EBADMSG";
			break;
		case EOVERFLOW:
			msg += "EOVERFLOW";
			break;
		// case ENOTUNIQ:
		// 	msg += "ENOTUNIQ";
		// 	break;
		// case EBADFD:
		// 	msg += "EBADFD";
		// 	break;
		// case EREMCHG:
		// 	msg += "EREMCHG";
		// 	break;
		// case ELIBACC:
		// 	msg += "ELIBACC";
		// 	break;
		// case ELIBBAD:
		// 	msg += "ELIBBAD";
		// 	break;
		// case ELIBSCN:
		// 	msg += "ELIBSCN";
		// 	break;
		// case ELIBMAX:
		// 	msg += "ELIBMAX";
		// 	break;
		// case ELIBEXEC:
		// 	msg += "ELIBEXEC";
		// 	break;
		case EILSEQ:
			msg += "EILSEQ";
			break;
		// case ERESTART:
		// 	msg += "ERESTART";
		// 	break;
		// case ESTRPIPE:
		// 	msg += "ESTRPIPE";
		// 	break;
		case EUSERS:
			msg += "EUSERS";
			break;
		case ENOTSOCK:
			msg += "ENOTSOCK";
			break;
		case EDESTADDRREQ:
			msg += "EDESTADDRREQ";
			break;
		case EMSGSIZE:
			msg += "EMSGSIZE";
			break;
		case EPROTOTYPE:
			msg += "EPROTOTYPE";
			break;
		case ENOPROTOOPT:
			msg += "ENOPROTOOPT";
			break;
		case EPROTONOSUPPORT:
			msg += "EPROTONOSUPPORT";
			break;
		case ESOCKTNOSUPPORT:
			msg += "ESOCKTNOSUPPORT";
			break;
		case EOPNOTSUPP:
			msg += "EOPNOTSUPP";
			break;
		case EPFNOSUPPORT:
			msg += "EPFNOSUPPORT";
			break;
		case EAFNOSUPPORT:
			msg += "EAFNOSUPPORT";
			break;
		case EADDRINUSE:
			msg += "EADDRINUSE";
			break;
		case EADDRNOTAVAIL:
			msg += "EADDRNOTAVAIL";
			break;
		case ENETDOWN:
			msg += "ENETDOWN";
			break;
		case ENETUNREACH:
			msg += "ENETUNREACH";
			break;
		case ENETRESET:
			msg += "ENETRESET";
			break;
		case ECONNABORTED:
			msg += "ECONNABORTED";
			break;
		case ECONNRESET:
			msg += "ECONNRESET";
			break;
		case ENOBUFS:
			msg += "ENOBUFS";
			break;
		case EISCONN:
			msg += "EISCONN";
			break;
		case ENOTCONN:
			msg += "ENOTCONN";
			break;
		case ESHUTDOWN:
			msg += "ESHUTDOWN";
			break;
		case ETOOMANYREFS:
			msg += "ETOOMANYREFS";
			break;
		case ETIMEDOUT:
			msg += "ETIMEDOUT";
			break;
		case ECONNREFUSED:
			msg += "ECONNREFUSED";
			break;
		case EHOSTDOWN:
			msg += "EHOSTDOWN";
			break;
		case EHOSTUNREACH:
			msg += "EHOSTUNREACH";
			break;
		case EALREADY:
			msg += "EALREADY";
			break;
		case EINPROGRESS:
			msg += "EINPROGRESS";
			break;
		case ESTALE:
			msg += "ESTALE";
			break;
		// case EUCLEAN:
		// 	msg += "EUCLEAN";
		// 	break;
		// case ENOTNAM:
		// 	msg += "ENOTNAM";
		// 	break;
		// case ENAVAIL:
		// 	msg += "ENAVAIL";
		// 	break;
		// case EISNAM:
		// 	msg += "EISNAM";
		// 	break;
		// case EREMOTEIO:
		// 	msg += "EREMOTEIO";
		// 	break;
		case EDQUOT:
			msg += "EDQUOT";
			break;
		// case ENOMEDIUM:
		// 	msg += "ENOMEDIUM";
		// 	break;
		// case EMEDIUMTYPE:
		// 	msg += "EMEDIUMTYPE";
		// 	break;
		case ECANCELED:
			msg += "ECANCELED";
			break;
		// case ENOKEY:
		// 	msg += "ENOKEY";
		// 	break;
		// case EKEYEXPIRED:
		// 	msg += "EKEYEXPIRED";
		// 	break;
		// case EKEYREVOKED:
		// 	msg += "EKEYREVOKED";
		// 	break;
		// case EKEYREJECTED:
		// 	msg += "EKEYREJECTED";
		// 	break;
		case EOWNERDEAD:
			msg += "EOWNERDEAD";
			break;
		case ENOTRECOVERABLE:
			msg += "ENOTRECOVERABLE";
			break;
		// case ERFKILL:
		// 	msg += "ERFKILL";
		// 	break;
		// case EHWPOISON:
		// 	msg += "EHWPOISON";
		// 	break;
		default:
			msg += to_string(en);
			break;
	}

	msg += ")";
}
#endif

mmap::mmap(fd_t h) {
#ifdef _WIN32
	LARGE_INTEGER fsli;

	if (!GetFileSizeEx(h, &fsli))
		throw last_error("GetFileSizeEx", GetLastError());

	filesize = fsli.QuadPart;

	mh = CreateFileMappingW(h, nullptr, PAGE_READONLY, 0, 0, nullptr);
	if (!mh)
		throw last_error("CreateFileMapping", GetLastError());
#else
	struct stat st;

	if (fstat(h, &st) == -1)
		throw errno_error("fstat", errno);

	filesize = st.st_size;

	ptr = ::mmap(nullptr, filesize, PROT_READ, MAP_SHARED, h, 0);

	if (ptr == MAP_FAILED)
		throw errno_error("mmap", errno);
#endif
}

mmap::~mmap() {
#ifdef _WIN32
	if (ptr)
		UnmapViewOfFile(ptr);

	if (mh)
		CloseHandle(mh);

	if (h != INVALID_HANDLE_VALUE)
		CloseHandle(h);
#else
	munmap(ptr, filesize);
#endif
}

span<const uint8_t> mmap::map() {
#ifdef _WIN32
	if (!ptr) {
		ptr = MapViewOfFile(mh, FILE_MAP_READ, 0, 0, 0);

		if (!ptr)
			throw last_error("MapViewOfFile", GetLastError());
	}
#endif

	return span((const uint8_t*)ptr, filesize);
}

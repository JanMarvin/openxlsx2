#pragma once

#include <string>

std::string b64encode(std::string_view sv);
std::string b64decode(std::string_view sv);

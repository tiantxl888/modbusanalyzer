#include <pybind11/pybind11.h>
#include <pybind11/stl.h>
#include <string>
#include <vector>
#include <map>

namespace py = pybind11;

class UpdaterPlugin {
public:
    UpdaterPlugin() : enabled(false) {}
    
    std::string get_name() const { return "Updater"; }
    std::string get_version() const { return "1.0.0"; }
    std::string get_description() const { return "Plugin for handling software updates"; }
    
    bool is_enabled() const { return enabled; }
    void set_enabled(bool value) { enabled = value; }
    
    std::map<std::string, std::string> get_config() const { return config; }
    void set_config(const std::map<std::string, std::string>& new_config) { config = new_config; }
    
private:
    bool enabled;
    std::map<std::string, std::string> config;
};

PYBIND11_MODULE(updater_plugin, m) {
    py::class_<UpdaterPlugin>(m, "UpdaterPlugin")
        .def(py::init<>())
        .def("get_name", &UpdaterPlugin::get_name)
        .def("get_version", &UpdaterPlugin::get_version)
        .def("get_description", &UpdaterPlugin::get_description)
        .def("is_enabled", &UpdaterPlugin::is_enabled)
        .def("set_enabled", &UpdaterPlugin::set_enabled)
        .def("get_config", &UpdaterPlugin::get_config)
        .def("set_config", &UpdaterPlugin::set_config);
} 
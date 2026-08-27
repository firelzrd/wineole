Gem::Specification.new do |spec|
  spec.name = 'wineole'
  spec.version = '0.1.0'
  spec.summary = 'Drive Win32OLE/COM automation of Wine-hosted Windows apps from Ruby'
  spec.description = <<~DESC
    A bridge (Rust, cross-compiled to run under Wine) plus a Ruby client
    that lets a Linux Ruby process drive Win32OLE/COM automation of
    Windows applications running under Wine, over a JSON Lines TCP
    protocol, without requiring a Windows build of Ruby.
  DESC
  spec.authors = ['firelzrd']
  spec.license = 'MIT'
  spec.homepage = 'https://github.com/firelzrd/wineole'
  spec.metadata = {
    'source_code_uri' => 'https://github.com/firelzrd/wineole',
    'bug_tracker_uri' => 'https://github.com/firelzrd/wineole/issues',
  }
  spec.required_ruby_version = '>= 3.0'
  spec.files = Dir['lib/**/*.rb'] + Dir['wineole-bridge-dist/**/*'] + Dir['LICENSE']
  spec.require_paths = ['lib']
end

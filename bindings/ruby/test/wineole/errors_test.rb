require 'minitest/autorun'
require_relative '../../lib/wineole/errors'

class ErrorsTest < Minitest::Test
  def test_remote_error_message_includes_class_and_message
    err = WineOLE::RemoteError.new('WIN32OLERuntimeError', 'boom')
    assert_equal 'WIN32OLERuntimeError: boom', err.message
    assert_equal 'WIN32OLERuntimeError', err.remote_class
  end

  def test_error_hierarchy
    assert WineOLE::NotSerializableError.ancestors.include?(WineOLE::Error)
    assert WineOLE::StaleReferenceError.ancestors.include?(WineOLE::Error)
    assert WineOLE::RemoteError.ancestors.include?(WineOLE::Error)
    assert WineOLE::ProtocolError.ancestors.include?(WineOLE::Error)
  end
end

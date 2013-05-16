package org.soqqo.tools.db;
import org.junit.Assert;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.test.context.ContextConfiguration;
import org.springframework.test.context.junit4.SpringJUnit4ClassRunner;

@RunWith(SpringJUnit4ClassRunner.class)
@ContextConfiguration( { "classpath:test-populator-config.xml" })
public class DBLoadingTest {

	@Autowired 
	private JdbcTemplate template;
	
	@Test
	public void testSingleSheet() { 
		// Sheet1 in src/test/resources/1sheet-sample.xls
		Assert.assertTrue(
				template.queryForInt("SELECT numberfield FROM table1 WHERE stringfield = 'foo 23'") == 38
				);
	}
}
